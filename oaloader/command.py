from pathlib import Path

import click
from tabulate import tabulate

from .const import (
    APP_NAME,
    DEFAULT_NETNAME,
    DEFAULT_PATH,
    OFFICE_SUBKEY_CATALOG,
    OFFICE_SUBKEY_PROVIDER,
    SUBKEY_OFFICE,
    version,
)
from .core import (
    add_manifests,
    clear_cache,
    enum_reg,
    fix_app_error,
    get_net_shares,
    load_manifest,
    local_server_url,
    office_installation,
    open_office_sub_key,
    remove_manifests,
    system_info,
)
from .log import setup_logger

__all__ = ["cli"]

opt_netname = click.option(
    "--netname",
    "-n",
    type=click.STRING,
    default=DEFAULT_NETNAME,
    show_default=True,
    help="NETNAME represents a share folder or a office add-in catalog.",
)
opt_path = click.option(
    "--path",
    "-p",
    type=click.Path(file_okay=False),
    default=DEFAULT_PATH,
    show_default=True,
    help="The directory stores manifests and is shared as NETNAME.",
)
arg_manifests = click.argument(
    "manifests",
    type=click.STRING,
    nargs=-1,
)


def use_decorators(*decorators):
    def wrapper(function):
        for d in reversed(decorators):
            function = d(function)
        return function

    return wrapper


@click.group()
@click.version_option(version=version, prog_name=APP_NAME)
@click.option(
    "--level",
    "-l",
    type=click.STRING,
    default="info",
    show_default=True,
    help="The log level",
)
def cli(level: str):
    """
    Manage your office add-ins locally.
    """
    setup_logger(level)


@cli.command()
@use_decorators(opt_netname, opt_path)
@click.option("--hide", is_flag=True, help="Hide add-ins in the catalog.")
@use_decorators(arg_manifests)
def add(netname, path, hide, manifests):
    """
    Register catalog and add manifests, manifests can be file paths or urls.

    NOTE: run as admin.
    """
    add_manifests(netname, path, hide, manifests)
    clear_cache()


@cli.command()
@use_decorators(opt_netname)
@click.option(
    "--all", "-a", "rm_all", is_flag=True, help="Remove all manifests in catalog path."
)
@click.option("--catalog", "-c", "rm_catalog", is_flag=True, help="Remove the catalog.")
@use_decorators(arg_manifests)
def remove(netname, rm_all, rm_catalog, manifests):
    """
    Remove manifest from catalog and manifest can be a file path or url.

    NOTE: run as admin.
    """
    remove_manifests(netname, rm_all, rm_catalog, manifests)
    clear_cache()


@cli.command()
def fix():
    """
    Try fixing `APP ERROR` when starting up add-ins.
    """
    fix_app_error()
    clear_cache()


def _echo_table(title, table_data):
    click.secho(title, bold=True, fg="green")
    if table_data:
        click.echo(tabulate(table_data, headers="keys"))
    else:
        click.secho("Nothing", fg="yellow")
    click.echo()


def _safe_open_key(title, sub_key):
    try:
        return open_office_sub_key(sub_key)
    except FileNotFoundError:
        click.secho(title, fg="green")
        raise click.ClickException(
            click.style(
                "Can't find Office 16 registry key, do you have a valid Microsoft installation?",
                fg="yellow",
            )
        )


@cli.command()
@use_decorators(opt_path)
def info(path):
    """
    Debug sideload status.
    """

    _echo_table("System Info:", [system_info()])
    _echo_table("Word Installation:", [office_installation("word")])

    net_shares = get_net_shares()
    _echo_table("Net Shares:", net_shares)

    title = rf"HKEY_CURRENT_USER\{SUBKEY_OFFICE}\{OFFICE_SUBKEY_PROVIDER}:"
    rv = filter(
        lambda v: v["attribute"] == "UniqueId",
        enum_reg(_safe_open_key(title, OFFICE_SUBKEY_PROVIDER)),
    )

    _echo_table(title, rv)

    title = rf"HKEY_CURRENT_USER\{SUBKEY_OFFICE}\{OFFICE_SUBKEY_CATALOG}:"
    rv = enum_reg(_safe_open_key(title, OFFICE_SUBKEY_CATALOG))
    _echo_table(title, rv)

    urls = []
    for item in rv:
        if item["attribute"] == "Url":
            urls.append(item["value"])

    url_to_path = {local_server_url(v["netname"]): v["path"] for v in net_shares}
    paths = set()
    paths.add(Path(path))
    for u in urls:
        p = url_to_path.get(u)
        if p:
            paths.add(Path(p))

    for path in paths:
        addins = []
        for p in path.glob("*.xml"):
            d, _ = load_manifest(str(p))
            addins.append(d)
        _echo_table(f"Office Add-ins in `{path}`:", addins)
