import pytest


def pytest_addoption(parser):
    parser.addoption(
        "--experimental",
        action="store_true",
        default=False,
        help="run experimental tests",
    )


def pytest_configure(config):
    config.addinivalue_line("markers", "experimental: mark test as experimental")


def pytest_collection_modifyitems(config, items):
    if not config.getoption("--experimental"):
        run_experimental = pytest.mark.skip(reason="need --experimental option to run")
        for item in items:
            if "experimental" in item.keywords:
                item.add_marker(run_experimental)
