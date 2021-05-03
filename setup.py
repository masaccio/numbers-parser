from setuptools import setup, find_packages

setup(
    name="numbers-parser",
    version="0.1",
    author="Jon Connell",
    author_email="python@figsandfudge.com",
    description="Package to read data from Apple Numbers spreadsheets",
    license="MIT",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/masaccio/numbers-parser",
    packages=find_packages(include=["numbers-parser", "numbers-parser.*"]),
    install_requires=[
        "binascii",
        "datetime",
        "google.protobuf",
        "plistlib",
        "pprint",
        "pytest",
        "snappy",
        "tqdm",
        "traceback",
        "yaml",
        "zipfile",
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)
