import getpass

import sentry_sdk

from oaloader import __version__, office_installation, system_info


def breadcrumb_sink(message):
    record = message.record
    sentry_sdk.add_breadcrumb(
        category=record["name"],
        level=record["level"].name,
        message=record["message"],
        data=record["extra"],
    )


def setup_sentry(sentry_dsn: str):
    sentry_sdk.init(
        sentry_dsn,
        send_default_pii=True,
        release=__version__,
    )
    sentry_sdk.set_user(
        dict(
            username=getpass.getuser(),
            ip_address="{{auto}}",
        )
    )

    info = system_info()
    sentry_sdk.set_context("System Info", info)
    sentry_sdk.set_tag("os.release", info["release"])
    sentry_sdk.set_tag("os.version", info["version"])

    office = office_installation("word")
    sentry_sdk.set_context("Word Installation", office)
    sentry_sdk.set_tag("office.version", office["version"])
    sentry_sdk.set_tag("office.build", office["build"])
