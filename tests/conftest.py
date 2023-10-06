import pytest


def pytest_addoption(parser):
    parser.addoption(
        "--experimental",
        action="store_true",
        default=False,
        help="run experimental tests",
    )
    parser.addoption("--save-file", action="store", default=None)
    parser.addoption(
        "--max-check-fails",
        default=False,
        type=int,
        help="maximum number of pytest.check failures",
    )


def pytest_configure(config):
    config.addinivalue_line("markers", "experimental: mark test as experimental")


def pytest_collection_modifyitems(config, items):
    if not config.getoption("--experimental"):
        run_experimental = pytest.mark.skip(reason="need --experimental option to run")
        for item in items:
            if "experimental" in item.keywords:
                item.add_marker(run_experimental)


@pytest.fixture(name="configurable_save_file")
def configurable_save_file_fixture(request, tmp_path, pytestconfig):
    if pytestconfig.getoption("save_file") is not None:
        new_filename = pytestconfig.getoption("save_file")
    else:
        new_filename = tmp_path / "test-save-new.numbers"

    yield new_filename


@pytest.fixture(name="configurable_multi_save_file", params="num_files")
def configurable_multi_save_file_fixture(request, tmp_path, pytestconfig):
    if isinstance(request.param, list) and len(request.param) == 1:
        num_files = request.param[0]
    else:
        num_files = 0

    if pytestconfig.getoption("save_file") is not None:
        new_filename = pytestconfig.getoption("save_file")
    else:
        new_filename = tmp_path / "test-save-new.numbers"

    if num_files <= 1:
        yield new_filename
    else:
        new_filenames = [str(new_filename).replace(".", f"-{x}.") for x in range(num_files)]
        yield new_filenames
