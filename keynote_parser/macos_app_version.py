class MacOSAppVersion(object):
    def __init__(self, short_version_string, bundle_version, build_version):
        self.short_version_string = short_version_string
        self.short_version_tuple = [int(x) for x in short_version_string.split('.')]
        self.bundle_version = bundle_version
        self.build_version = build_version

    def __str__(self):
        return "%s (%s, %s)" % (self.short_version_string, self.bundle_version, self.build_version)

    def __lt__(self, value):
        if not isinstance(value, MacOSAppVersion):
            raise TypeError("< not supported between MacOSAppVersion and %s" % type(value))
        return (
            self.short_version_comparator < value.short_version_comparator
            or self.bundle_version < value.bundle_version
            or self.build_version < value.build_version
        )

    def __le__(self, value):
        if not isinstance(value, MacOSAppVersion):
            raise TypeError("<= not supported between MacOSAppVersion and %s" % type(value))
        return (
            self.short_version_comparator <= value.short_version_comparator
            or self.bundle_version <= value.bundle_version
            or self.build_version <= value.build_version
        )

    @property
    def major(self):
        return self.short_version_tuple[0]

    @property
    def minor(self):
        return self.short_version_tuple[1]

    @property
    def short_version_comparator(self):
        return sum(
            [
                (10 ** (len(self.short_version_tuple) - i)) * value
                for i, value in enumerate(self.short_version_tuple)
            ]
        )
