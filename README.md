# Office Addin Sideloader

[![PyPI](https://img.shields.io/pypi/v/oaloader?style=flat-square)](https://pypi.org/project/oaloader/)
![GitHub](https://img.shields.io/github/license/elonzh/office-addin-sideloader?style=flat-square)

A handy tool to manage your office addins locally,
you can use it for addin development or deploy your addins for your clients out of AppSource.

> NOTE: currently only support windows.

## Features

- Add or remove Office Addin locally.
- Support local or url manifest source.
- Debug sideload status and list manifest info.
- Single binary without any dependency.
- Use it as a library.
- Generate addin installer by single command.

## Installation

### Pre-built releases

If you just use the command line and don't have a python environment,
download pre-built binary from [GitHub Releases](https://github.com/elonzh/office-addin-sideloader/releases).

### Pypi

```shell
> pip install oaloader
```

## Quick Start

```shell
> ./oaloader.exe --help
Usage: oaloader.exe [OPTIONS] COMMAND [ARGS]...

  Manage your office addins locally.

Options:
  --version         Show the version and exit.
  -l, --level TEXT  The log level  [default: info]
  --help            Show this message and exit.

Commands:
  add     Register catalog and add manifests, manifests can be file paths
          or...
  info    Debug sideload status.
  remove  Remove manifest from catalog and manifest can be a file path or...
```

## Build an Addin installer

1. Install [Poetry](https://python-poetry.org/docs/).
2. Run `poetry install` to prepare environment.
3. Run `poetry run invoke installer -m <YOUR-ADDIN-MANIFEST-URL>` to build your own installer.

If your want customize the installer, just edit `installer.jinja2` or write your own installer with `oaloader` module.

## How it works

https://docs.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins
