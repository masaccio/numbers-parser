import logging
from enum import Flag, auto

logger = logging.getLogger(__name__)
debug = logger.debug


class ExperimentalFeatures(Flag):
    NONE = auto()
    TESTING = auto()
    GROUPED_CATEGORIES = auto()


EXPERIMENTAL_FEATURES = ExperimentalFeatures.NONE


def enable_experimental_feature(flags: ExperimentalFeatures) -> None:
    global EXPERIMENTAL_FEATURES
    EXPERIMENTAL_FEATURES |= flags
    debug("Experimental features: enabling %s, flags=%s", flags, EXPERIMENTAL_FEATURES)


def disable_experimental_feature(flags: ExperimentalFeatures) -> None:
    global EXPERIMENTAL_FEATURES
    EXPERIMENTAL_FEATURES ^= flags
    debug("Experimental features: disabling %s, flags=%s", flags, EXPERIMENTAL_FEATURES)


def experimental_features() -> ExperimentalFeatures:
    return EXPERIMENTAL_FEATURES
