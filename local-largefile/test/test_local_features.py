"""Run with OFFICE_EXTENSION_BASE pointing to the unmodified 4.2.3 package."""
import os
from pathlib import Path
import re
import subprocess
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from local_features import patch_host


class LocalFeaturesTest(unittest.TestCase):
    def test_incompatible_bundle_is_rejected(self):
        with self.assertRaises(ValueError):
            patch_host('unrecognized bundle')

    @unittest.skipUnless(os.environ.get('OFFICE_EXTENSION_BASE'), 'requires official 4.2.3 bundle')
    def test_desktop_and_web_host_without_license_or_network(self):
        for web in (False, True):
            with self.subTest(web=web):
                name = 'extension.web.js' if web else 'extension.js'
                original = (Path(os.environ['OFFICE_EXTENSION_BASE']) / 'out' / name).read_text()
                patched = patch_host(original, web)
                predicate, cancelled, gate, reader = (
                    ('Dw', '_f', 'BN', 'Ds') if web else ('Kte', 'U_', 'iG', 'zA'))
                functions = []
                for function in (predicate, cancelled, gate):
                    functions.append(re.search(
                        rf'function {function}\([^)]*\)\{{[^}}]*\}}', patched).group())
                # Run the actual patched host predicates with a cancelled account.
                # They must not read storage, emit a paywall event or use a license key.
                probe = '\n'.join(functions) + f'''
                  const assert = require('node:assert/strict');
                  const context = {{globalState: {{get() {{throw Error('storage accessed')}}}}}};
                  const handler = {{emit() {{throw Error('paywall emitted')}}}};
                  assert.equal({predicate}(context), true);
                  assert.equal({cancelled}(context), false);
                  assert.equal({gate}(context, handler), true);
                '''
                subprocess.run(['node', '-e', probe], check=True)
                self.assertIn(f'static getLicenseKey(e){{return {reader}(e)}}', patched)
                self.assertNotRegex(patched, r'\w+\.refreshLatestLicense\(t\)')
                # Rebuilding a patched host must fail instead of silently patching it twice.
                with self.assertRaises(ValueError):
                    patch_host(patched, web)


if __name__ == '__main__':
    unittest.main()
