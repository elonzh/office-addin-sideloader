from pathlib import Path

import click

version = "0.4.1"

SUBKEY_CATALOG = r"Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
SUBKEY_PROVIDER = r"Software\Microsoft\Office\16.0\WEF\Providers"
APP_NAME = "office-addin-sideloader"
APP_DIR = Path(click.get_app_dir(APP_NAME))
MANIFEST_ENCODING = "utf-8"
DEFAULT_PATH = APP_DIR.joinpath("addins")
DEFAULT_NETNAME = "office-addins"
