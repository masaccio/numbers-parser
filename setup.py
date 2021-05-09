from setuptools import setup, find_packages

setup(
    name="numbers-parser",
    version="1.0.2",
    author="Jon Connell",
    author_email="python@figsandfudge.com",
    description="Package to read data from Apple Numbers spreadsheets",
    license="MIT",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/masaccio/numbers-parser",
    package_dir={'': 'src'},
    packages=find_packages(where='src'),
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
    python_requires='>=3.6',
)
