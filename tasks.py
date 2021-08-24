import hashlib
import re
import sys
from pathlib import Path

import jinja2
from invoke import Context, Result, task

nuitka_base_args = [
    sys.executable,
    "-m",
    "nuitka",
    "--assume-yes-for-downloads",
    "--onefile",
    "--output-dir=build",
    "--remove-output",
    # Windows Defender will report Trojan:Win32/Sabsik.TE.A!ml
    # when using --nofollow-imports
    # "--nofollow-imports",
]


def rename_executable(output: str, executable: str):
    matched = re.search(r"Successfully created '(.+)'\.", output)
    if not matched:
        return
    src = Path(matched.group(1))
    dst = src.with_name(executable + src.suffix)
    dst.unlink(missing_ok=True)
    dst = src.rename(dst)
    print(f"Rename executable to {dst}")
    return dst


@task()
def build(c, executable="oaloader"):
    """
    Build oaloader.
    """
    c: Context

    args = nuitka_base_args.copy()
    args.extend(["oaloader/__main__.py"])
    r: Result = c.run(" ".join(args))
    rename_executable(r.stdout, executable)


def make_executable(
    c,
    template,
    manifests,
    success_msg,
    fail_msg,
    sentry_dsn="",
    executable="addin-executable",
    icon="",
    dry=False,
):
    c: Context
    env = jinja2.Environment()
    env.filters["repr"] = repr
    t = env.from_string(Path(template).read_text())
    script = t.render(
        manifests=manifests,
        success_msg=success_msg,
        fail_msg=fail_msg,
        sentry_dsn=sentry_dsn,
    )

    p = Path(f"executable_{hashlib.md5(script.encode()).hexdigest()[:8]}").with_suffix(
        ".py"
    )
    p.write_text(script, encoding="utf-8")
    args = nuitka_base_args.copy()
    args.extend(
        [
            "--include-package=sentry_sdk",
            "--windows-uac-admin",
            # "--windows-disable-console",
        ]
    )
    if icon:
        args.extend(
            [
                "--windows-icon-from-ico",
                icon,
            ]
        )
    args.append(str(p))
    r: Result = c.run(" ".join(args), dry=dry)
    if dry:
        print(f"Dry run mode, delete file {p!s} manually")
        return executable
    else:
        e = rename_executable(r.stdout, executable)
        p.unlink(missing_ok=True)
        return e


@task(iterable=["manifests"])
def installer(
    c,
    manifests,
    success_msg="Addins are successfully installed. Press any key to continue...",
    fail_msg="Operation failed, do you have a valid Microsoft installation?",
    sentry_dsn="",
    executable="addin-installer",
    icon="",
    template="bin/installer.jinja2",
    dry=False,
):
    """
    Build an addin installer.
    """
    return make_executable(
        c,
        template=template,
        manifests=manifests,
        success_msg=success_msg,
        fail_msg=fail_msg,
        sentry_dsn=sentry_dsn,
        executable=executable,
        icon=icon,
        dry=dry,
    )


@task(iterable=["manifests"])
def uninstaller(
    c,
    manifests,
    success_msg="Addins are successfully uninstalled. Press any key to continue...",
    fail_msg="Operation failed, do you have a valid Microsoft installation?",
    sentry_dsn="",
    executable="addin-uninstaller",
    icon="",
    template="bin/uninstaller.jinja2",
    dry=False,
):
    """
    Build an addin uninstaller.
    """
    return make_executable(
        c,
        template=template,
        manifests=manifests,
        success_msg=success_msg,
        fail_msg=fail_msg,
        sentry_dsn=sentry_dsn,
        executable=executable,
        icon=icon,
        dry=dry,
    )
