from setuptools import setup, find_packages
from distutils.util import convert_path

main_ns = {}
ver_path = convert_path('src/numbers_parser/_version.py')
with open(ver_path) as ver_file:
    exec(ver_file.read(), main_ns)

setup(
    name="numbers-parser",
    version=main_ns['__version__'],
    author="Jon Connell",
    author_email="python@figsandfudge.com",
    description="Package to read data from Apple Numbers spreadsheets",
    license="MIT",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/masaccio/numbers-parser",
    package_dir={"": "src"},
    packages=find_packages(where="src"),
    entry_points = {
        'console_scripts': [
            'cat-numbers=numbers_parser._cat_numbers:main',
            'unpack-numbers=numbers_parser._unpack_numbers:main'
        ],
    },
    install_requires=[
        "protobuf",
        "pytest",
        "python-snappy",
        "PyYAML",
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.6",
)
