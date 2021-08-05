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
    "--output-dir",
    "build",
    # Windows Defender will report Trojan:Win32/Sabsik.TE.A!ml
    # when using --nofollow-imports
    # "--nofollow-imports",
    "--follow-imports",
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
    c: Context

    args = nuitka_base_args.copy()
    args.extend(["oaloader/__main__.py"])
    r: Result = c.run(" ".join(args))
    rename_executable(r.stdout, executable)


@task(iterable=["manifests"])
def installer(
    c,
    manifests,
    msg="Addins are successfully installed. Press any key to continue...",
    executable="addin-installer",
    icon="",
    template="installer.jinja2",
):
    c: Context
    env = jinja2.Environment()
    env.filters["repr"] = repr
    t = env.from_string(Path(template).read_text())
    script = t.render(manifests=manifests, msg=msg)

    p = Path(f"installer_{hashlib.md5(script.encode()).hexdigest()[:8]}").with_suffix(
        ".py"
    )
    p.write_text(script, encoding="utf-8")
    args = nuitka_base_args.copy()
    args.extend(
        [
            "--windows-uac-admin",
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
    r: Result = c.run(" ".join(args))
    rename_executable(r.stdout, executable)
    p.unlink(missing_ok=True)
