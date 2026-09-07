"""Offline regression gates for generated compiler metadata and reproducible skill archives."""
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile

from bundle import build, validate_archive

ROOT = Path(__file__).resolve().parent
COMMANDER = ROOT / 'm365-surface-commander' / 'scripts'
sys.path.insert(0, str(COMMANDER))
from surface_cli import analyze
from surface_cli.normalizer import normalize


class ManifestContractTests(unittest.TestCase):
    def run_isolated(self, skill, parser, manifest, mutation):
        with tempfile.TemporaryDirectory() as temporary:
            scripts = Path(temporary) / 'scripts'
            shutil.copytree(ROOT / skill / 'scripts', scripts, ignore=shutil.ignore_patterns('__pycache__'))
            path = scripts / manifest
            mutation(path)
            result = subprocess.run([sys.executable, str(scripts / parser)], input='```cmd\n/unknown value=1\n```',
                                    capture_output=True, text=True, check=False)
            self.assertNotEqual(result.returncode, 0, result.stdout)
            self.assertIn('Rebuild with skills:generate', result.stderr)
            self.assertEqual(result.stdout, '')

    def test_missing_corrupt_or_incompatible_manifest_never_enables_fallback(self):
        for skill, parser, filename in (
            ('m365-surface-commander', 'parse_commands.py', 'm365-cli-1.0.json'),
            ('m365-command-planner', 'parse_plan.py', 'm365-plan-1.0.json'),
        ):
            for mutation in (
                lambda path: path.unlink(), lambda path: path.write_text('{bad'),
                lambda path: path.write_text(json.dumps({'version': 'future/99'})),
            ):
                with self.subTest(skill=skill, mutation=mutation):
                    self.run_isolated(skill, parser, filename, mutation)

    def test_incomplete_authority_metadata_is_a_packaging_error(self):
        def mutate(path):
            data = json.loads(path.read_text())
            data['preflight']['approvalByKind'].pop('set-reaction')
            path.write_text(json.dumps(data))
        self.run_isolated('m365-surface-commander', 'parse_commands.py', 'm365-cli-1.0.json', mutate)

    def test_unsupported_guard_metadata_never_weakens_validation(self):
        for replacement in (
            {},
            {'query': {'requiredColumns': {'type': 'string', 'checks': [{'kind': 'regex'}]}}},
            {'query': {'requiredColumns': {'type': 'number', 'checks': [{'kind': 'min', 'value': float('nan')}]}}},
        ):
            def mutate(path):
                data = json.loads(path.read_text())
                data['preflight']['analysisGuards'] = replacement
                path.write_text(json.dumps(data))
            with self.subTest(replacement=replacement):
                self.run_isolated('m365-surface-commander', 'parse_commands.py', 'm365-cli-1.0.json', mutate)

    def test_read_phase_and_effect_risk_follow_contract_metadata(self):
        for line in ('ls /doc', 'find /doc -name *.md', 'tail /work/report.txt'):
            result = analyze(line)
            self.assertEqual(result['errors'], [])
            self.assertEqual(result['reads'], [line])
            self.assertEqual(result['effects'], [])
            self.assertEqual(normalize('set A1 1\n' + line)[0][0], line)
        for line, authority, reversible in (
            ('/set-reaction messageId=1 reactionType=like', 'estate', True),
            ('/delete-message messageId=1', 'irreversible', False),
            ('share evidence.txt = "text"', 'estate', False),
        ):
            effect = analyze(line)['effects'][0]
            self.assertEqual(effect['approvalClass'], authority)
            self.assertEqual(effect['reversible'], reversible)
            self.assertTrue(effect['external'])
        materialize = analyze('analyze {"kind":"materialize","input":"artifact:fixture","destination":"A1"}')['effects'][0]
        self.assertEqual(materialize['approvalClass'], 'in-document')
        self.assertTrue(materialize['reversible'])
        self.assertFalse(materialize['external'])

    def test_archives_are_content_deterministic_and_validate_current_source(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            skill = root / 'fixture'
            skill.mkdir()
            source = skill / 'SKILL.md'
            source.write_text('---\nname: fixture\ndescription: Fixture\n---\n')
            first, second = root / 'first.zip', root / 'second.zip'
            build(skill, first)
            os.utime(source, (1000000000, 1000000000))
            build(skill, second)
            self.assertEqual(hashlib.sha256(first.read_bytes()).digest(), hashlib.sha256(second.read_bytes()).digest())
            self.assertEqual(validate_archive(skill, first), [])
            self.assertTrue(validate_archive(skill, root / 'missing.zip'))
            source.write_text(source.read_text() + 'Changed instructions\n')
            self.assertTrue(validate_archive(skill, first))
            build(skill, first)
            with zipfile.ZipFile(first, 'a') as archive:
                archive.writestr('../outside.txt', 'bad')
            self.assertTrue(validate_archive(skill, first))

    def test_bundle_rejects_links_and_excludes_interpreter_caches(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            (root / 'SKILL.md').write_text('fixture')
            scripts = root / 'scripts'
            scripts.mkdir()
            (scripts / 'cache.pyc').write_bytes(b'not-source')
            (scripts / 'outside.py').symlink_to(root / 'SKILL.md')
            with self.assertRaises(ValueError):
                build(root, root / 'fixture.zip')
            (scripts / 'outside.py').unlink()
            build(root, root / 'fixture.zip')
            with zipfile.ZipFile(root / 'fixture.zip') as archive:
                self.assertEqual(archive.namelist(), ['SKILL.md'])


if __name__ == '__main__':
    unittest.main()
