"""Numbers cache module helpers."""

from __future__ import annotations

from collections import defaultdict
from functools import wraps
from typing import TYPE_CHECKING, Any

from typing_extensions import Self

if TYPE_CHECKING:
    from collections.abc import Callable


class Cacheable:
    """Provide per-instance storage for cached method results."""

    _cache: defaultdict[str, Any]

    def __new__(
        cls,
        *_args: object,
        **_kwargs: object,
    ) -> Self:
        """Create an instance and initialize its cache."""
        obj = object.__new__(cls)
        obj._cache = defaultdict(lambda: defaultdict(dict))
        return obj


def cache(
    num_args: int = 1,
) -> Callable[[Callable[..., object]], Callable[..., object]]:
    """
    Cache a method using a selected subset of its arguments.

    The returned decorator memoizes results for each instance and argument key.
    """

    def cache_decorator(func: Callable[..., object]) -> Callable[..., object]:
        """Wrap a method with argument-aware cache lookups."""

        @wraps(func)
        def inner_multi_args(
            self: Cacheable,
            *args: object,
            **kwargs: object,
        ) -> object:
            """Return a cached result keyed by the selected arguments."""
            method = func.__name__
            key = ".".join([str(args[x]) for x in range(num_args)])
            if key in self._cache[method]:
                return self._cache[method][key]
            value = func(self, *args, **kwargs)
            self._cache[method][key] = value
            return value

        def inner_no_args(self: Cacheable) -> object:
            """Return a cached result for a method without arguments."""
            method = func.__name__
            if method not in self._cache:
                self._cache[method] = func(self)
            return self._cache[method]

        if num_args == 0:
            return inner_no_args
        return inner_multi_args

    return cache_decorator
