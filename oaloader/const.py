from pathlib import Path

import click

version = "0.4.4"

SUBKEY_OFFICE = r"Software\Microsoft\Office\16.0"
OFFICE_SUBKEY_CATALOG = r"WEF\TrustedCatalogs"
OFFICE_SUBKEY_PROVIDER = r"WEF\Providers"
APP_NAME = "office-addin-sideloader"
APP_DIR = Path(click.get_app_dir(APP_NAME))
MANIFEST_ENCODING = "utf-8"
DEFAULT_PATH = APP_DIR.joinpath("addins")
DEFAULT_NETNAME = "office-addins"
