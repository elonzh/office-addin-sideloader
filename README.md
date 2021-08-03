# Office Addin Sideloader

A handy tool to manage your office addins locally,
you can use it for addin development or deploy your addins for your clients out of AppSource

> NOTE: currently only support windows.

## Features

- Add or remove Office Addin locally.
- Debug sideload status and list manifest info.
- Single binary without any dependency.

## Quick Start

```shell
> ./oaloader.exe --help
Usage: oaloader.exe [OPTIONS] COMMAND [ARGS]...

  manage your office addins locally.

Options:
  --version         Show the version and exit.
  -l, --level TEXT  log level
  --help            Show this message and exit.

Commands:
  add     add manifest to catalog.
  info    debug sideload status.
  remove  remove manifest from catalog.
```

## How it works

https://docs.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins
