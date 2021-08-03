import os
import shutil
import sys
import winreg
from pathlib import Path

import click
import pywintypes
import win32net
import win32netcon
import xmltodict
from loguru import logger
from tabulate import tabulate

server = None  # Run on local machine.
subkey = r"Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"

__version__ = "0.0.1"
app_name = "office-addin-sideloader"
app_dir = Path(click.get_app_dir(app_name))

default_path = app_dir.joinpath("addins")
default_netname = "office-addins"


def setup_logger(level: str):
    level = level.upper()
    logger.remove()
    fmt = (
        "<green>{time:YYYY-MM-DD HH:mm:ss.SSS}</green> | "
        "<level>{level: <8}</level> | "
        "<level>{message}</level> - "
        "{extra}"
    )
    logger.add(
        sys.stdout,
        level=level,
        format=fmt,
    )
    logger.add(
        app_dir.joinpath("main.log"),
        level="DEBUG",
        format=fmt,
    )


def local_server_url(netname):
    local_server = win32net.NetServerGetInfo(None, 100)
    url = rf"\\{local_server['name'].lower()}\{netname}"
    return url


def get_net_shares(share_type=win32netcon.STYPE_DISKTREE):
    net_shares, _, _ = win32net.NetShareEnum(server, 2)
    if share_type is not None:
        net_shares = list(filter(lambda v: v["type"] == share_type, net_shares))
    return net_shares


def add_net_share(netname: str, path) -> str:
    """
    https://docs.microsoft.com/en-us/windows/win32/api/lmshare/nf-lmshare-netshareadd
    """
    path = os.path.abspath(path)
    share_info = {
        "netname": netname,
        "path": path,
        "type": win32netcon.STYPE_DISKTREE,
        "permissions": win32netcon.ACCESS_READ,
        "remark": "",
        "max_uses": -1,
        "current_uses": 0,
        "passwd": "",
    }
    url = local_server_url(netname)
    context_logger = logger.bind(netname=netname, url=url, path=path)
    try:
        win32net.NetShareAdd(server, 2, share_info)
        context_logger.debug("add net share")
    except win32net.error as e:
        #  (2118, 'NetShareAdd', '名称已经共享。')
        if e.winerror != 2118:
            raise
        context_logger.debug("net share already exists")
    return url


def remove_net_share(netname: str):
    rv = win32net.NetShareDel(server, netname)
    logger.bind(netname=netname).debug("remove net share")
    return rv


def find_catalog(root, url: str):
    for item in enum_reg(root):
        if item["attribute"] == "Url" and item["value"] == url:
            _id = item["key"]
            return _id
    logger.bind(url=url).debug("can not find catalog")


def add_catalog(root, url: str, hide: bool = False):
    _id = find_catalog(root, url) or str(pywintypes.CreateGuid())
    key = winreg.CreateKey(root, _id)
    winreg.SetValueEx(key, "Id", 0, winreg.REG_SZ, _id)
    winreg.SetValueEx(key, "Url", 0, winreg.REG_SZ, url)
    # show in menu
    winreg.SetValueEx(
        key,
        "Flags",
        0,
        winreg.REG_DWORD,
        int(not hide),
    )
    logger.bind(id=_id, url=url, hide=hide).debug("add catalog")
    return key


def remove_catalog(root, url: str):
    _id = find_catalog(root, url)
    if _id:
        winreg.DeleteKey(root, _id)
        logger.bind(key=_id, url=url).debug("remove catalog")


def enum_reg(root):
    keys = []
    i = 0
    while True:
        try:
            keys.append(winreg.EnumKey(root, i))
            i += 1
        except winreg.error as e:
            # [WinError 259] 没有可用的数据了。
            if e.winerror == 259:
                break
            else:
                raise

    rv = []
    for key in keys:
        k = winreg.OpenKey(root, key)
        i = 0
        while True:
            try:
                a, v, t = winreg.EnumValue(k, i)
                rv.append(
                    {
                        "key": key,
                        "attribute": a,
                        "value": v,
                        "type": t,
                    }
                )
                i += 1
            except winreg.error as e:
                # [WinError 259] 没有可用的数据了。
                if e.winerror == 259:
                    break
                else:
                    raise
    return rv


def load_manifest(path):
    path = Path(path)
    manifest = xmltodict.parse(path.read_text())["OfficeApp"]
    return {
        "file": path.name,
        "id": manifest["Id"],  # must have
        "version": manifest.get("Version"),
        "provider_name": manifest.get("ProviderName"),
        "display_name": manifest.get("DisplayName", {}).get("@DefaultValue"),
        "description": manifest.get("Description", {}).get("@DefaultValue"),
    }


opt_netname = click.option(
    "--netname", "-n", type=click.STRING, default=default_netname
)
opt_path = click.option(
    "--path", "-p", type=click.Path(file_okay=False), default=default_path
)


def use_options(*options):
    def wrapper(function):
        for option in reversed(options):
            function = option(function)
        return function

    return wrapper


@click.group()
@click.version_option(version=__version__, prog_name=app_name)
@click.option("--level", "-l", type=click.STRING, default="info", help="log level")
def cli(level: str):
    """
    manage your office addins locally.
    """
    setup_logger(level)


@cli.command()
@use_options(opt_netname, opt_path)
@click.option("--hide", is_flag=True, help="hide addins in the catalog")
@click.argument(
    "manifests", type=click.Path(exists=True, file_okay=True, dir_okay=False), nargs=-1
)
def add(netname, path, hide, manifests):
    """
    add manifest to catalog.

    NOTE: run as admin.
    """
    path = Path(path)
    if not path.exists():
        path.mkdir(parents=True, exist_ok=True)
    url = add_net_share(netname, path)
    root = winreg.OpenKey(winreg.HKEY_CURRENT_USER, subkey)
    add_catalog(root, url, hide)

    for m in manifests:
        d = load_manifest(m)
        dst = path.joinpath(f"{d['id']}.xml")
        shutil.copy(m, dst)
        logger.bind(src=m, dst=str(dst)).debug("copy manifest")


@cli.command()
@use_options(opt_netname)
@click.option("--all", "-a", "rm_all", is_flag=True, help="remove all manifests")
@click.option("--catalog", "-c", "rm_catalog", is_flag=True, help="remove the catalog")
@click.argument(
    "manifests", type=click.Path(exists=True, file_okay=True, dir_okay=False), nargs=-1
)
def remove(netname, rm_all, rm_catalog, manifests):
    """
    remove manifest from catalog.

    NOTE: run as admin.
    """
    try:
        share = win32net.NetShareGetInfo(server, netname, 2)
    except win32net.error as e:
        # 2310, 'NetShareGetInfo', '共享资源不存在。'
        if e.winerror == 2310:
            logger.info("net share does not exist")
            return
        raise
    path = Path(share["path"])

    if rm_all:
        manifests = path.glob("*.xml")
    for m in manifests:
        d = load_manifest(m)
        dst = path.joinpath(f"{d['id']}.xml")
        dst.unlink(missing_ok=True)
        logger.bind(src=m, dst=str(dst)).debug("remove manifest")

    if rm_catalog:
        root = winreg.OpenKey(winreg.HKEY_CURRENT_USER, subkey)
        url = local_server_url(netname)
        remove_catalog(root, url)
        remove_net_share(netname)


def _echo_table(title, table_data):
    click.secho(title, bold=True, fg="green")
    if table_data:
        click.echo(tabulate(table_data, headers="keys"))
    else:
        click.secho("Nothing", fg="yellow")
    click.echo()


@cli.command()
@use_options(opt_path)
def info(path):
    """
    debug sideload status.
    """
    root = winreg.OpenKey(winreg.HKEY_CURRENT_USER, subkey)
    rv = enum_reg(root)
    _echo_table(rf"HKEY_CURRENT_USER\{subkey}:", rv)

    urls = []
    for item in rv:
        if item["attribute"] == "Url":
            urls.append(item["value"])

    net_shares = get_net_shares()

    _echo_table("Net Shares:", net_shares)

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
            addins.append(load_manifest(p))
        _echo_table(f"Office Addins in `{path}`:", addins)


if __name__ == "__main__":
    cli()
