import getpass

import sentry_sdk
from sentry_sdk.integrations.atexit import AtexitIntegration
from sentry_sdk.integrations.dedupe import DedupeIntegration
from sentry_sdk.integrations.excepthook import ExcepthookIntegration

from oaloader import office_installation, system_info


def breadcrumb_sink(message):
    record = message.record
    sentry_sdk.add_breadcrumb(
        category=record["name"],
        level=record["level"].name,
        message=record["message"],
        extra=record["extra"],
    )


def setup_sentry(sentry_dsn: str):
    sentry_sdk.init(
        sentry_dsn,
        send_default_pii=True,
        auto_enabling_integrations=False,
        integrations=(
            ExcepthookIntegration(),
            DedupeIntegration(),
            AtexitIntegration(),
        ),
    )
    sentry_sdk.set_user(
        dict(
            username=getpass.getuser(),
            ip_address="{{auto}}",
        )
    )
    sentry_sdk.set_context("System Info", system_info())
    sentry_sdk.set_context("Word Installation", office_installation("word"))
