from functools import wraps
from collections import defaultdict


class Cacheable:
    def __new__(cls, *args, **kwargs):
        obj = object.__new__(cls)
        obj._cache = defaultdict(lambda: defaultdict(dict))
        return obj


def cache(num_args=1):
    """Decorator to memoize a class method using a precise subset of
    its arguments"""

    def cache_decorator(func):
        @wraps(func)
        def inner_multi_args(self, *args, **kwargs):
            method = func.__name__
            key = ".".join([str(args[x]) for x in range(num_args)])
            if key in self._cache[method]:
                return self._cache[method][key]
            else:
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
        else:
            return inner_multi_args

    return cache_decorator
