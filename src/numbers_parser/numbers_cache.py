from collections import defaultdict
from functools import wraps


class Cacheable:
    def __new__(cls, *args, **kwargs):
        obj = object.__new__(cls)
        obj._cache = defaultdict(lambda: defaultdict(dict))
        return obj


def cache(num_args=1):
    """
    Decorator to memoize a class method using a precise subset of its
    arguments. The cache is stored in the `_cache` attribute of the instance
    ensuring that all caches are instance-specific rather than class-specific.

    Args:
        num_args (int): The number of arguments to use as the cache key. If set to 0,
                        the method will be cached without any arguments.

    Returns:
        function: The decorated function with caching applied.

    Usage:
        @cache(num_args=2)
        def some_method(self, arg1, arg2):
            # Method implementation
            pass

    """

    def cache_decorator(func):
        @wraps(func)
        def inner_multi_args(self, *args, **kwargs):
            method = func.__name__
            key = ".".join([str(args[x]) for x in range(num_args)])
            if key in self._cache[method]:
                return self._cache[method][key]
            value = func(self, *args, **kwargs)
            self._cache[method][key] = value
            return value

        def inner_no_args(self):
            method = func.__name__
            if method not in self._cache:
                self._cache[method] = func(self)
            return self._cache[method]

        if num_args == 0:
            return inner_no_args
        return inner_multi_args

    return cache_decorator
