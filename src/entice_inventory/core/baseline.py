"""Select a native MARIO reader without changing parsing or accounting rules."""
from pathlib import Path

from entice_inventory.core.paths import data_path, local_config


def load_baseline(path=None, input_format=None):
    import mario

    path = Path(path) if path is not None else data_path('baseline', required=True)
    input_format = input_format or local_config().get('baseline_format', 'mario_parquet')
    if input_format == 'mario_parquet':
        return mario.parse_from_parquet(str(path), table='IOT', mode='flows')
    if input_format in {'gtap_csv', 'gtap_gdx'}:
        if path.suffix.lower() == '.zip':
            raise ValueError('Extract the original GTAP bundle first and select its directory.')
        return mario.parse_gtap(str(path), table='IOT', variant='power',
                                layout='MRIO', input_format=input_format.removeprefix('gtap_'),
                                year=2023)
    raise ValueError('baseline_format must be mario_parquet, gtap_csv or gtap_gdx')
