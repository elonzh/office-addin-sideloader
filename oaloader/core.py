import os
import platform
import re
import shutil
import winreg
from pathlib import Path
from urllib.request import urlopen

import pywintypes
import win32com.client
import win32net
import win32netcon
import xmltodict
from loguru import logger

from . import const

__all__ = [
    "local_server_url",
    "get_net_shares",
    "add_net_share",
    "remove_net_share",
    "find_catalog",
    "add_catalog",
    "remove_catalog",
    "enum_reg",
    "load_manifest",
    "add_manifests",
    "remove_manifests",
    "system_info",
    "office_installation",
    "fix_app_error",
    "clear_cache",
]


def local_server_url(netname):
    local_server = win32net.NetServerGetInfo(None, 100)
    url = rf"\\{local_server['name'].lower()}\{netname}"
    return url


def get_net_shares(share_type=win32netcon.STYPE_DISKTREE):
    net_shares, _, _ = win32net.NetShareEnum(None, 2)
    if share_type is not None:
        net_shares = list(filter(lambda v: v["type"] == share_type, net_shares))
    return net_shares


def add_net_share(netname: str, path) -> str:
    """
    https://docs.microsoft.com/en-us/windows/win32/api/lmshare/nf-lmshare-netshareadd
    """
    path = Path(path).absolute()
    if not path.exists():
        path.mkdir(parents=True, exist_ok=True)

    path_str = str(path)
    share_info = {
        "netname": netname,
        "path": path_str,
        "type": win32netcon.STYPE_DISKTREE,
        "permissions": win32netcon.ACCESS_READ,
        "remark": "",
        "max_uses": -1,
        "current_uses": 0,
        "passwd": "",
    }
    url = local_server_url(netname)
    context_logger = logger.bind(netname=netname, url=url, path=path_str)
    try:
        win32net.NetShareAdd(None, 2, share_info)
        context_logger.debug("add net share")
    except win32net.error as e:
        #  (2118, 'NetShareAdd', '名称已经共享。')
        if e.winerror != 2118:
            raise
        context_logger.debug("net share already exists")
    return url


def remove_net_share(netname: str):
    rv = win32net.NetShareDel(None, netname)
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

    for key in keys:
        k = winreg.OpenKey(root, key)
        i = 0
        while True:
            try:
                a, v, t = winreg.EnumValue(k, i)
                yield {
                    "key": key,
                    "attribute": a,
                    "value": v,
                    "type": t,
                }
                i += 1
            except winreg.error as e:
                # [WinError 259] 没有可用的数据了。
                if e.winerror == 259:
                    break
                else:
                    raise


url_pattern = re.compile(
    r"https?://(www.)?[-a-zA-Z0-9@:%._+~#=]{1,256}.[a-zA-Z0-9()]{1,6}\b([-a-zA-Z0-9()!@:%_+.~#?&/=]*)"
)


def is_url(s: str) -> bool:
    return bool(url_pattern.match(s))


def load_manifest(src: str):
    if is_url(src):
        logger.bind(url=src).debug("load manifest from url")
        raw = urlopen(src).read().decode(const.MANIFEST_ENCODING)
    else:
        raw = Path(src).read_text(encoding=const.MANIFEST_ENCODING)
    data = xmltodict.parse(raw)["OfficeApp"]
    return {
        "id": data["Id"],  # must have
        "version": data.get("Version"),
        "provider_name": data.get("ProviderName"),
        "display_name": data.get("DisplayName", {}).get("@DefaultValue"),
        "description": data.get("Description", {}).get("@DefaultValue"),
    }, raw


def add_manifests(
    netname=const.DEFAULT_NETNAME,
    path=const.DEFAULT_PATH,
    hide=False,
    manifests=tuple(),
):
    """
    Register catalog and add manifests, manifests can be file paths or urls.

    NOTE: run as admin.
    """
    # open register key first make sure the office installation exists
    root = winreg.OpenKey(winreg.HKEY_CURRENT_USER, const.SUBKEY_CATALOG)
    url = add_net_share(netname, path)
    add_catalog(root, url, hide)

    for m in manifests:
        d, raw = load_manifest(m)
        dst = path.joinpath(f"{d['id']}.xml")
        dst.write_text(raw, encoding=const.MANIFEST_ENCODING)
        logger.bind(src=m, dst=str(dst)).debug("add manifest")


def remove_manifests(
    netname=const.DEFAULT_NETNAME, rm_all=False, rm_catalog=False, manifests=tuple()
):
    """
    Remove manifest from catalog and manifest can be a file path or url.

    NOTE: run as admin.
    """
    try:
        share = win32net.NetShareGetInfo(None, netname, 2)
    except win32net.error as e:
        # 2310, 'NetShareGetInfo', '共享资源不存在。'
        if e.winerror == 2310:
            logger.bind(netname=netname).info("net share does not exist")
            return
        raise
    path = Path(share["path"])

    if rm_all:
        manifests = path.glob("*.xml")
    for m in manifests:
        d, _ = load_manifest(str(m))
        dst = path.joinpath(f"{d['id']}.xml")
        dst.unlink(missing_ok=True)
        logger.bind(src=m, dst=str(dst)).debug("remove manifest")

    if rm_catalog:
        root = winreg.OpenKey(winreg.HKEY_CURRENT_USER, const.SUBKEY_CATALOG)
        url = local_server_url(netname)
        remove_catalog(root, url)
        remove_net_share(netname)


def fix_app_error():
    """
    https://docs.microsoft.com/en-us/office365/troubleshoot/installation/cannot-install-office-add-in
    """
    try:
        root = winreg.OpenKey(winreg.HKEY_CURRENT_USER, const.SUBKEY_PROVIDER)
    except FileNotFoundError:
        logger.bind(subykey=const.SUBKEY_PROVIDER).debug("subkey not found")
        return

    for v in enum_reg(root):
        if v["attribute"] == "UniqueId" and (
            v["value"] == "Anonymous" or v["value"].endswith("_ADAL")
        ):
            k = rf"{const.SUBKEY_PROVIDER}\{v['key']}"
            logger.bind(subykey=k, attribute=v["attribute"], value=v["value"]).debug(
                "delete invalid key"
            )
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, k)


def clear_cache():
    """
    https://docs.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache
    """
    paths = [
        Path(os.getenv("LOCALAPPDATA")).joinpath(r"Microsoft\Office\16.0\Wef"),
        Path(os.getenv("USERPROFILE")).joinpath(
            r"AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache"
        ),
    ]

    def onerror(func, path, exc_info):
        log = logger.bind(func=func, path=path, exc_info=exc_info)
        level = "WARNING"
        if exc_info[0] == FileNotFoundError:
            level = "DEBUG"
        log.log(level, "rmtree failed")

    for p in paths:
        logger.bind(path=str(p)).debug("rmtree")
        shutil.rmtree(
            p,
            onerror=onerror,
        )


def system_info():
    return dict(
        platform=platform.platform(),
        system=platform.system(),
        release=platform.release(),
        version=platform.version(),
    )


def office_installation(app: str):
    try:
        word = win32com.client.Dispatch(f"{app.capitalize()}.Application")
        # some attrs may not exist
        return dict(
            name=getattr(word, "Name", "Unknown"),
            version=getattr(word, "Version", "Unknown"),
            build=getattr(word, "Build", "Unknown"),
            path=getattr(word, "Path", "Unknown"),
            startup_path=getattr(word, "StartupPath", "Unknown"),
        )
    except pywintypes.com_error as e:
        # com_error: (-2147221005, '无效的类字符串', None, None)
        # com_error: (-2146959355, '服务器运行失败', None, None)
        if e.hresult in (-2147221005, -2146959355):
            return dict(
                name="Unknown",
                version="Unknown",
                build="Unknown",
                path="Unknown",
                startup_path="Unknown",
            )
        raise
