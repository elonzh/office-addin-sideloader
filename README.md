# Office Addin Sideloader

[![PyPI](https://img.shields.io/pypi/v/oaloader?style=flat-square)](https://pypi.org/project/oaloader/)
![GitHub](https://img.shields.io/github/license/elonzh/office-addin-sideloader?style=flat-square)

A handy tool to manage your office addins locally,
you can use it for addin development or deploy your addins for your clients out of AppSource.

> NOTE: currently only support windows.

## Features

- Add or remove Office Add-in locally.
- Support local or url manifest source.
- Debug sideload status and list manifest info.
- Single binary without any dependency.
- Use it as a library.
- Generate add-in installer/uninstaller with [sentry](https://sentry.io) support by single command.
- Support fixing add-in [APP ERROR](https://docs.microsoft.com/en-us/office365/troubleshoot/installation/cannot-install-office-add-in) and [clearing cache](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache).

## Installation

### Pre-built releases

If you just use the command line and don't have a python environment,
download pre-built binary from [GitHub Releases](https://github.com/elonzh/office-addin-sideloader/releases).

### Pypi

```shell
> pip install oaloader
```

## Quick Start

```text
> ./oaloader.exe --help
Usage:  [OPTIONS] COMMAND [ARGS]...

  Manage your office addins locally.

Options:
  --version         Show the version and exit.
  -l, --level TEXT  The log level  [default: info]
  --help            Show this message and exit.

Commands:
  add     Register catalog and add manifests, manifests can be file paths
          or...

  fix     Try fixing `APP ERROR` when starting up add-ins.
  info    Debug sideload status.
  remove  Remove manifest from catalog and manifest can be a file path or...
```

## Build an Addin installer

1. Install [Poetry](https://python-poetry.org/docs/).
2. Run `poetry install` to prepare environment.
3. Run `poetry run invoke installer -m <YOUR-ADDIN-MANIFEST-URL>` to build your own installer.

If your want customize the installer, just edit `installer.jinja2` or write your own installer with `oaloader` module.

## Build an Addin uninstaller

Just using invoke `uninstaller` task like `installer` above.

## How it works

https://docs.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins
