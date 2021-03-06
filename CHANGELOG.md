# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.4.4] - 2021-10-22
### Fixed
- Office subkey may not exist.

## [0.4.3] - 2021-10-13
### Fixed
- Fix potential ssl issue when using urlopen.

## [0.4.2] - 2021-09-08
### Changed
- Exception in `fix_app_error` and `clear_cache` will be captured silently.
### Fixed
- Fix setting sentry breadcrumb data wrongly.

## [0.4.1] - 2021-08-26
### Changed
- `setup_sentry` now `set_tag` about office installation and use default integrations.
- installer/uninstaller shows version now.
### Fixed
- Fix `fix_app_error` deletes wrong keys.

## [0.4.0] - 2021-08-24
### Added
- Support fixing add-in and clear cache.
- Support build uninstaller.

### Fixed
- Fix office_installation error.
- Update Nuitka to develop version to fix sign issue.
- more robust installer/uninstaller behavior.

## [0.3.0] - 2021-08-12
### Added
- Support check office installation.
- Add dry run mode for installer task.

## [0.2.1] - 2021-08-12
### Changed
- Return unknown when find office installation failed.

## [0.2.0] - 2021-08-09
### Added
- Add sentry_sdk support for installer.
- Support debug system info and office installation.

## [0.1.0] - 2021-08-05
### Added
- Release to pypi.
