"""A storage/format choice must not alter MARIO's numerical settings."""
import sys
import unittest
from unittest.mock import Mock, patch

from entice_inventory.core.baseline import load_baseline


class BaselineReaderTests(unittest.TestCase):
    def test_parquet_replay_retains_native_flow_reader(self):
        mario = Mock()
        with patch.dict(sys.modules, {'mario': mario}):
            result = load_baseline('/cache/gtap', 'mario_parquet')
        mario.parse_from_parquet.assert_called_once_with('/cache/gtap', table='IOT', mode='flows')
        mario.parse_gtap.assert_not_called()
        self.assertIs(result, mario.parse_from_parquet.return_value)

    def test_original_formats_use_native_gtap_reader_without_fallback(self):
        for fmt in ('csv', 'gdx'):
            with self.subTest(fmt=fmt):
                mario = Mock()
                with patch.dict(sys.modules, {'mario': mario}):
                    load_baseline('/original/gtap', 'gtap_'+fmt)
                mario.parse_gtap.assert_called_once_with('/original/gtap', table='IOT',
                    variant='power', layout='MRIO', input_format=fmt, year=2023)
                mario.parse_from_parquet.assert_not_called()

    def test_wrong_format_or_unextracted_zip_cannot_fall_back_to_cache(self):
        mario = Mock()
        with patch.dict(sys.modules, {'mario': mario}):
            for path, fmt in [('/x','unknown'),('/original/gtap.zip','gtap_csv')]:
                with self.assertRaises(ValueError):load_baseline(path, fmt)
        mario.parse_gtap.assert_not_called()
        mario.parse_from_parquet.assert_not_called()
