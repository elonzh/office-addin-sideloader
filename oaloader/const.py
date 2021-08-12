from pathlib import Path

import click

version = "0.2.1"
server = None  # Run on local machine.
subkey = r"Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
app_name = "office-addin-sideloader"
app_dir = Path(click.get_app_dir(app_name))
manifest_encoding = "utf-8"
default_path = app_dir.joinpath("addins")
default_netname = "office-addins"
