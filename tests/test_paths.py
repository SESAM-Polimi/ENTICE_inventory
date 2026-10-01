"""Shared data lookup must not silently select a personal cloud directory."""
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from entice_inventory.core import paths


class SharedPathsTests(unittest.TestCase):
    def test_config_and_environment_precedence(self):
        with tempfile.TemporaryDirectory() as tmp, patch.dict(os.environ, {}, clear=True):
            root=Path(tmp).resolve()
            config=root/'paths.local.json'
            config.write_text(json.dumps({'data_root':'group-data'}))
            with patch.object(paths,'PROJECT_ROOT',root):
                self.assertEqual(paths.project_data_root(),root/'group-data')
                with patch.dict(os.environ,{'ENTICE_DATA_ROOT':str(root/'override')}):
                    self.assertEqual(paths.project_data_root(),root/'override')

    def test_external_config_and_licensed_totals_lookup(self):
        with tempfile.TemporaryDirectory() as tmp, patch.dict(os.environ, {}, clear=True):
            root=Path(tmp).resolve();shared=root/'shared';inputs=shared/'Repository inputs';inputs.mkdir(parents=True)
            totals=inputs/'GTAP12_X.xlsx';totals.write_bytes(b'lookup-only fixture')
            config=root/'custom.json';config.write_text(json.dumps({'data_root':str(shared)}))
            with patch.dict(os.environ,{'ENTICE_CONFIG':str(config)}), patch.object(paths,'DATA_DIR',root/'checkout-data'):
                self.assertEqual(paths.project_data_root(required=True),shared)
                self.assertEqual(paths.find_data_file('GTAP12_X.xlsx'),totals)
                explicit=root/'explicit';explicit.mkdir();(explicit/'GTAP12_X.xlsx').write_bytes(b'explicit')
                self.assertEqual(paths.find_data_file('GTAP12_X.xlsx',explicit),explicit/'GTAP12_X.xlsx')

    def test_missing_and_invalid_configuration_are_explicit(self):
        with tempfile.TemporaryDirectory() as tmp, patch.dict(os.environ, {}, clear=True):
            root=Path(tmp).resolve()
            with patch.object(paths,'PROJECT_ROOT',root):
                with self.assertRaises(FileNotFoundError):paths.project_data_root(required=True)
                self.assertEqual(paths.project_data_root(),root/'external-data')
                (root/'paths.local.json').write_text('{"data_root": []}')
                with self.assertRaises(ValueError):paths.project_data_root()


if __name__=='__main__':
    unittest.main()
