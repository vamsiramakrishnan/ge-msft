#!/usr/bin/env python3
"""
Offline, dependency-free unit tests for the skill provisioning tooling (create_skill.py,
test_skill.py).

These tests cover Finding #8 of the external review: the tooling shipped with live
project/engine *defaults* and destructive --replace/--share with no confirmation. The hardening
contract these tests pin:

  1) No real PROJECT/PROJECT_NUMBER/ENGINE defaults — a LIVE operation REQUIRES GE_PROJECT,
     GE_PROJECT_NUMBER, GE_ENGINE in the environment, else the tool refuses with a clear message.
  2) DRY-RUN is the default; --replace and --share (destructive / tenant-wide) require an explicit
     --yes. A dry run performs ZERO network calls.
  3) Every request carries a timeout and bounded retries, and calls raise_for_status().

The tests must run with NO live credentials and NO third-party deps beyond the stdlib (google.auth
in particular is import-lazy in the tooling so this file can import it offline). Run:

    python3 skill/test_tooling.py
"""
import importlib.util
import base64
import json
import sys
import unittest
from pathlib import Path
from unittest import mock

HERE = Path(__file__).parent


def _load(module_name: str, filename: str):
    """Import a sibling script by path WITHOUT triggering its __main__ block."""
    spec = importlib.util.spec_from_file_location(module_name, HERE / filename)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = mod
    spec.loader.exec_module(mod)
    return mod


# Import lazily-dependent modules. These must NOT pull in google.auth / requests at import time.
create_skill = _load("create_skill", "create_skill.py")
update_skills = _load("update_skills", "update_skills.py")


def _fake_jwt(payload):
    raw = json.dumps(payload, separators=(",", ":")).encode("utf-8")
    encoded = base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")
    return f"header.{encoded}.signature"


class _FakeResponse:
    """Minimal requests.Response stand-in that records raise_for_status() calls."""

    def __init__(self, status_code=200, payload=None, headers=None, content=None):
        self.status_code = status_code
        self._payload = payload or {}
        self.headers = headers or {}
        self.content = b"{}" if content is None else content
        self.text = self.content.decode("utf-8", errors="replace")
        self.raised = False

    def raise_for_status(self):
        self.raised = True
        if self.status_code >= 400:
            raise RuntimeError(f"HTTP {self.status_code}")

    def json(self):
        return self._payload


class _RecordingSession:
    """Records every HTTP call so tests can assert dry-run does zero network I/O and that timeouts
    are always supplied."""

    def __init__(self, payload=None):
        self.calls = []
        self._payload = payload or {}
        self.headers = {}

    def _record(self, verb, url, **kw):
        self.calls.append((verb, url, kw))
        assert "timeout" in kw and kw["timeout"], f"{verb} {url} missing timeout"
        return _FakeResponse(200, self._payload)

    def get(self, url, **kw):
        return self._record("GET", url, **kw)

    def post(self, url, **kw):
        return self._record("POST", url, **kw)

    def patch(self, url, **kw):
        return self._record("PATCH", url, **kw)

    def delete(self, url, **kw):
        return self._record("DELETE", url, **kw)


class _SequenceSession(_RecordingSession):
    def __init__(self, responses):
        super().__init__()
        self._responses = list(responses)

    def _record(self, verb, url, **kw):
        self.calls.append((verb, url, kw))
        assert "timeout" in kw and kw["timeout"], f"{verb} {url} missing timeout"
        if not self._responses:
            raise AssertionError(f"unexpected HTTP call {verb} {url}")
        return self._responses.pop(0)


# ---------------------------------------------------------------------------
# 1) Live config is required — no real defaults.
# ---------------------------------------------------------------------------
class TestLiveConfigRequired(unittest.TestCase):
    REQUIRED = ("GE_PROJECT", "GE_PROJECT_NUMBER", "GE_ENGINE")

    def test_no_real_defaults_baked_in(self):
        """The module source must not carry the reviewer-flagged live project/engine identifiers."""
        src = (HERE / "create_skill.py").read_text()
        for needle in ("vital-octagon-19612", "440790012685", "phoenix-telco_1751440131886"):
            self.assertNotIn(needle, src, f"live default {needle!r} still baked into create_skill.py")

    def test_resolve_config_refuses_when_unset(self):
        with mock.patch.dict("os.environ", {}, clear=True):
            with self.assertRaises(SystemExit) as cm:
                create_skill.resolve_live_config()
            msg = str(cm.exception)
            for var in self.REQUIRED:
                self.assertIn(var, msg, f"refusal message should name {var}")

    def test_resolve_config_refuses_on_partial(self):
        partial = {"GE_PROJECT": "p", "GE_PROJECT_NUMBER": "123"}  # ENGINE missing
        with mock.patch.dict("os.environ", partial, clear=True):
            with self.assertRaises(SystemExit) as cm:
                create_skill.resolve_live_config()
            self.assertIn("GE_ENGINE", str(cm.exception))

    def test_resolve_config_ok_when_all_set(self):
        env = {"GE_PROJECT": "p", "GE_PROJECT_NUMBER": "123", "GE_ENGINE": "e"}
        with mock.patch.dict("os.environ", env, clear=True):
            cfg = create_skill.resolve_live_config()
        self.assertEqual(cfg.project, "p")
        self.assertEqual(cfg.project_number, "123")
        self.assertEqual(cfg.engine, "e")
        self.assertIn("123", cfg.assistant)
        self.assertIn("e", cfg.assistant)


# ---------------------------------------------------------------------------
# 2) Dry-run is default; --replace/--share need --yes; dry-run does no network I/O.
# ---------------------------------------------------------------------------
class TestDryRunDefault(unittest.TestCase):
    ENV = {
        "GE_PROJECT": "p",
        "GE_PROJECT_NUMBER": "123",
        "GE_ENGINE": "e",
        "GE_WIDGET_CONFIG_ID": "33333333-3333-4333-8333-333333333333",
    }

    def _run(self, argv):
        with mock.patch.dict("os.environ", self.ENV, clear=True):
            with mock.patch.object(create_skill, "session") as sess_factory:
                rec = _RecordingSession({"name": "agents/x", "displayName": "X"})
                sess_factory.return_value = rec
                # ensure a zip path exists so we exercise the create/upload path
                zip_path = HERE / "m365-surface-commander.zip"
                created_zip = False
                if not zip_path.exists():
                    zip_path.write_bytes(b"PK\x03\x04stub")
                    created_zip = True
                try:
                    code = create_skill.main(argv)
                finally:
                    if created_zip:
                        zip_path.unlink()
                return code, rec, sess_factory

    def test_dry_run_is_default_and_makes_no_network_calls(self):
        code, rec, sess_factory = self._run([])
        self.assertEqual(code, 0)
        # In dry-run the session factory should never even be called.
        sess_factory.assert_not_called()
        self.assertEqual(rec.calls, [], "dry-run must perform zero HTTP calls")

    def test_public_list_dry_run_makes_no_network_calls(self):
        code, rec, sess_factory = self._run(["--api-mode", "public", "--list"])
        self.assertEqual(code, 0)
        sess_factory.assert_not_called()
        self.assertEqual(rec.calls, [], "public --list dry-run must perform zero HTTP calls")

    def test_replace_without_yes_refuses(self):
        with mock.patch.dict("os.environ", self.ENV, clear=True):
            with self.assertRaises(SystemExit) as cm:
                create_skill.main(["--live", "--replace"])
            self.assertIn("--yes", str(cm.exception))

    def test_share_without_yes_refuses(self):
        with mock.patch.dict("os.environ", self.ENV, clear=True):
            with self.assertRaises(SystemExit) as cm:
                create_skill.main(["--live", "--share"])
            self.assertIn("--yes", str(cm.exception))

    def test_live_without_config_refuses(self):
        with mock.patch.dict("os.environ", {}, clear=True):
            with self.assertRaises(SystemExit) as cm:
                create_skill.main(["--live"])
            self.assertIn("GE_PROJECT", str(cm.exception))


class TestBatchSkillUpdater(unittest.TestCase):
    ENV = {
        "GE_PROJECT": "p",
        "GE_PROJECT_NUMBER": "123",
        "GE_ENGINE": "e",
        "GE_WIDGET_CONFIG_ID": "33333333-3333-4333-8333-333333333333",
    }

    def test_batch_dry_run_makes_no_network_calls(self):
        with mock.patch.dict("os.environ", self.ENV, clear=True):
            with mock.patch.object(create_skill, "session") as sess_factory:
                code = update_skills.main([])
        self.assertEqual(code, 0)
        sess_factory.assert_not_called()

    def test_batch_replace_requires_yes(self):
        with mock.patch.dict("os.environ", self.ENV, clear=True):
            with self.assertRaises(SystemExit) as cm:
                update_skills.main(["--replace"])
            self.assertIn("--yes", str(cm.exception))

    def test_batch_delete_only_requires_yes(self):
        with mock.patch.dict("os.environ", self.ENV, clear=True):
            with self.assertRaises(SystemExit) as cm:
                update_skills.main(["--delete-only"])
            self.assertIn("--yes", str(cm.exception))

    def test_batch_widget_live_requires_explicit_existing_or_create_new(self):
        with mock.patch.dict("os.environ", self.ENV, clear=True):
            with self.assertRaises(SystemExit) as cm:
                update_skills.main(["--live"])
            self.assertIn("--upload-existing", str(cm.exception))


# ---------------------------------------------------------------------------
# 3) Requests carry timeouts + raise_for_status; retry wrapper is bounded.
# ---------------------------------------------------------------------------
class TestRequestHardening(unittest.TestCase):
    ENV = {"GE_PROJECT": "p", "GE_PROJECT_NUMBER": "123", "GE_ENGINE": "e"}

    def test_create_shell_uses_timeout_and_raises(self):
        with mock.patch.dict("os.environ", self.ENV, clear=True):
            cfg = create_skill.resolve_live_config()
            rec = _RecordingSession({"name": "agents/x"})
            create_skill.create_shell(rec, cfg, "agent-x", "instruction")
        self.assertEqual(len(rec.calls), 1)
        verb, url, kw = rec.calls[0]
        self.assertEqual(verb, "POST")
        self.assertIn("timeout", kw)

    def test_delete_raises_for_status(self):
        with mock.patch.dict("os.environ", self.ENV, clear=True):
            cfg = create_skill.resolve_live_config()
            rec = _RecordingSession()
            resp = create_skill.delete_agent(rec, cfg, "agent-x")
        self.assertTrue(resp.raised, "delete must call raise_for_status()")

    def test_retrying_session_is_bounded(self):
        # The HTTP adapter must mount a bounded Retry (total is a small finite int, not infinite).
        retries = create_skill.MAX_RETRIES
        self.assertIsInstance(retries, int)
        self.assertGreaterEqual(retries, 1)
        self.assertLessEqual(retries, 10)
        self.assertIsInstance(create_skill.HTTP_TIMEOUT, (int, float))
        self.assertGreater(create_skill.HTTP_TIMEOUT, 0)

    def test_widget_session_rejects_gcloud_adc_token_shape(self):
        env = {
            **self.ENV,
            "GE_WIDGET_BEARER_TOKEN": "ya29.local-adc-token",
        }
        with mock.patch.dict("os.environ", env, clear=True):
            cfg = create_skill.resolve_live_config()
            with self.assertRaises(SystemExit) as cm:
                create_skill.session(cfg, "widget")
        self.assertIn("not a Vertex AI Search widget JWT", str(cm.exception))

    def test_widget_session_accepts_widget_jwt_shape(self):
        token = _fake_jwt(
            {
                "iss": "https://vertexaisearch.cloud.google",
                "aud": "https://content-discoveryengine.googleapis.com",
                "exp": 4102444800,
            }
        )
        env = {
            **self.ENV,
            "GE_WIDGET_BEARER_TOKEN": token,
            "GE_WIDGET_SERVER_TOKEN": "CAMSAh0H",
        }
        with mock.patch.dict("os.environ", env, clear=True):
            cfg = create_skill.resolve_live_config()
            s = create_skill.session(cfg, "widget")
        self.assertEqual(s.headers["Authorization"], f"Bearer {token}")
        self.assertEqual(s.headers["x-server-token"], "CAMSAh0H")

    def test_list_agents_uses_public_agents_endpoint(self):
        with mock.patch.dict("os.environ", self.ENV, clear=True):
            cfg = create_skill.resolve_live_config()
            rec = _RecordingSession({"agents": [{"name": "agents/x"}]})
            result = create_skill.list_agents(rec, cfg)
        self.assertEqual(result["agents"][0]["name"], "agents/x")
        verb, url, kw = rec.calls[0]
        self.assertEqual(verb, "GET")
        self.assertTrue(url.endswith("/agents"))
        self.assertIn("discoveryengine.googleapis.com", url)
        self.assertIn("timeout", kw)

    def test_widget_create_uses_content_api_and_returns_numeric_agent_name(self):
        env = {
            **self.ENV,
            "GE_WIDGET_CONFIG_ID": "33333333-3333-4333-8333-333333333333",
        }
        with mock.patch.dict("os.environ", env, clear=True):
            cfg = create_skill.resolve_live_config()
            widget = create_skill.resolve_widget_config()
            rec = _RecordingSession({"agent": {"name": "8870098647237058037"}})
            agent = create_skill.create_widget_agent(
                rec, cfg, widget, "instruction", "Name", "Desc"
            )
        self.assertEqual(agent["name"], "8870098647237058037")
        verb, url, kw = rec.calls[0]
        self.assertEqual(verb, "POST")
        self.assertIn("content-discoveryengine.googleapis.com", url)
        self.assertIn("widgetCreateAgent", url)
        self.assertEqual(kw["json"]["configId"], env["GE_WIDGET_CONFIG_ID"])
        self.assertTrue(kw["json"]["createAgentRequest"]["defaultFilesSkipped"])
        self.assertNotIn("params", kw)

    def test_widget_delete_uses_content_api_agent_name(self):
        env = {
            **self.ENV,
            "GE_WIDGET_CONFIG_ID": "33333333-3333-4333-8333-333333333333",
        }
        with mock.patch.dict("os.environ", env, clear=True):
            cfg = create_skill.resolve_live_config()
            widget = create_skill.resolve_widget_config()
            rec = _SequenceSession(
                [
                    _FakeResponse(
                        status_code=400,
                        content=(
                            b'{"error":{"message":"Invalid JSON payload received. Unknown name '
                            b'\\"deleteAgentRequest\\": Cannot find field."}}'
                        ),
                    ),
                    _FakeResponse(payload={}),
                ]
            )
            resp = create_skill.delete_widget_agent(rec, cfg, widget, "8870098647237058037")
        self.assertEqual(resp.status_code, 200)
        self.assertEqual(len(rec.calls), 2)
        verb, url, kw = rec.calls[1]
        self.assertEqual(verb, "POST")
        self.assertIn("content-discoveryengine.googleapis.com", url)
        self.assertIn("widgetDeleteAgent", url)
        self.assertEqual(kw["json"]["name"], "8870098647237058037")

    def test_widget_zip_upload_uses_resumable_protocol(self):
        with mock.patch.dict("os.environ", self.ENV, clear=True):
            cfg = create_skill.resolve_live_config()
        zip_path = HERE / "m365-surface-commander.zip"
        rec = _SequenceSession(
            [
                _FakeResponse(
                    headers={
                        "x-goog-upload-url": "https://content-discoveryengine.googleapis.com/upload/session"
                    },
                    content=b"",
                ),
                _FakeResponse(payload={}, content=b"{}"),
            ]
        )
        create_skill.upload_zip_resumable(rec, cfg, "8870098647237058037", zip_path)
        self.assertEqual(len(rec.calls), 2)
        start = rec.calls[0]
        final = rec.calls[1]
        self.assertIn("/agents/8870098647237058037/files:upload", start[1])
        self.assertEqual(start[2]["headers"]["x-goog-upload-command"], "start")
        self.assertEqual(start[2]["headers"]["x-goog-upload-protocol"], "resumable")
        self.assertEqual(
            start[2]["headers"]["x-goog-upload-header-content-length"],
            str(zip_path.stat().st_size),
        )
        self.assertEqual(final[1], "https://content-discoveryengine.googleapis.com/upload/session")
        self.assertEqual(final[2]["headers"]["x-goog-upload-command"], "upload, finalize")
        self.assertEqual(final[2]["headers"]["x-goog-upload-offset"], "0")


# ---------------------------------------------------------------------------
# test_skill.py: live mode must also refuse without config; stub stays offline.
# ---------------------------------------------------------------------------
class TestTestSkillLiveGuard(unittest.TestCase):
    def test_source_has_no_live_defaults(self):
        src = (HERE / "test_skill.py").read_text()
        for needle in ("vital-octagon-19612", "440790012685", "phoenix-telco_1751440131886"):
            self.assertNotIn(needle, src, f"live default {needle!r} still baked into test_skill.py")


# ---------------------------------------------------------------------------
# surface_cli.py: the deterministic preflight compiler (ADR-0008 §4). Its --self-test
# covers check/budget/plan, capability scope, dependency inference, and risk.
# ---------------------------------------------------------------------------
class TestSurfaceCliPreflight(unittest.TestCase):
    def test_self_test_passes(self):
        import subprocess

        script = HERE / "m365-surface-commander" / "scripts" / "surface_cli.py"
        proc = subprocess.run(
            [sys.executable, str(script), "--self-test"], capture_output=True, text=True
        )
        self.assertEqual(proc.returncode, 0, f"surface_cli self-test failed:\n{proc.stderr}")

    def test_eval_recall_and_no_false_positives(self):
        # ADR-0008 §10 — the offline eval gate: surface_cli must catch every seeded defect (recall
        # 100%) and never flag a valid program. Guards the helper's quality as the language evolves.
        import subprocess

        harness = HERE / "eval" / "eval_harness.py"
        proc = subprocess.run([sys.executable, str(harness)], capture_output=True, text=True)
        self.assertEqual(proc.returncode, 0, f"eval gate failed:\n{proc.stdout}\n{proc.stderr}")


if __name__ == "__main__":
    unittest.main(verbosity=2)
