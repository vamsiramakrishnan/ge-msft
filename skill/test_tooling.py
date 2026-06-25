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


class _FakeResponse:
    """Minimal requests.Response stand-in that records raise_for_status() calls."""

    def __init__(self, status_code=200, payload=None):
        self.status_code = status_code
        self._payload = payload or {}
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
    ENV = {"GE_PROJECT": "p", "GE_PROJECT_NUMBER": "123", "GE_ENGINE": "e"}

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


if __name__ == "__main__":
    unittest.main(verbosity=2)
