"""Experimental module helpers."""

import logging
from enum import Flag, auto

logger = logging.getLogger(__name__)
debug = logger.debug


class ExperimentalFeatures(Flag):
    """Manage feature flags that enable experimental behavior."""

    NONE = auto()
    TESTING = auto()
    GROUPED_CATEGORIES = auto()


EXPERIMENTAL_FEATURES = ExperimentalFeatures.NONE


def enable_experimental_feature(flags: ExperimentalFeatures) -> None:
    """Enable experimental feature."""
    global EXPERIMENTAL_FEATURES  # noqa: PLW0603
    EXPERIMENTAL_FEATURES |= flags
    debug("Experimental features: enabling %s, flags=%s", flags, EXPERIMENTAL_FEATURES)


def disable_experimental_feature(flags: ExperimentalFeatures) -> None:
    """Disable experimental feature."""
    global EXPERIMENTAL_FEATURES  # noqa: PLW0603
    EXPERIMENTAL_FEATURES ^= flags
    debug("Experimental features: disabling %s, flags=%s", flags, EXPERIMENTAL_FEATURES)


def experimental_features() -> ExperimentalFeatures:
    """Return the names of available experimental features."""
    return EXPERIMENTAL_FEATURES
