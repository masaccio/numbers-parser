import logging

logger = logging.getLogger(__name__)
debug = logger.debug

_EXPERIMENTAL_FEATURES = False


def _enable_experimental_features(status: bool) -> None:
    global _EXPERIMENTAL_FEATURES
    _EXPERIMENTAL_FEATURES = status
    debug("Experimental features %s", "on" if status else "off")


def _experimental_features() -> bool:
    return _EXPERIMENTAL_FEATURES
