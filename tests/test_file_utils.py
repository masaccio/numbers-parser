from numbers_parser import codec, file_utils

NUMBERS_FILE = "tests/data/sofware-comparison.numbers"
FILE_LIST = [
    "Data/ImageFillFromTexture-15.tiff",
    "Data/ImageFillFromTexture-16.tiff",
    "Data/ImageFillFromTexture-17.tiff",
    "Data/ImageFillFromTexture-18.tiff",
    "Data/ImageFillFromTexture-19.tiff",
    "Index/AnnotationAuthorStorage.iwa",
    "Index/CalculationEngine.iwa",
    "Index/Document.iwa",
    "Index/DocumentMetadata.iwa",
    "Index/DocumentStylesheet.iwa",
    "Index/Metadata.iwa",
    "Index/ObjectContainer.iwa",
    "Index/Tables/DataList-874955.iwa",
    "Index/Tables/DataList-874956.iwa",
    "Index/Tables/DataList-874957.iwa",
    "Index/Tables/DataList-874958.iwa",
    "Index/Tables/DataList-874959.iwa",
    "Index/Tables/DataList-874960.iwa",
    "Index/Tables/DataList-874961.iwa",
    "Index/Tables/DataList-874962.iwa",
    "Index/Tables/DataList-874963.iwa",
    "Index/Tables/DataList-874964.iwa",
    "Index/Tables/DataList-874965.iwa",
    "Index/Tables/DataList-874984.iwa",
    "Index/Tables/DataList-874985.iwa",
    "Index/Tables/DataList-874986.iwa",
    "Index/Tables/DataList-874987.iwa",
    "Index/Tables/DataList-874988.iwa",
    "Index/Tables/DataList-874989.iwa",
    "Index/Tables/DataList-874990.iwa",
    "Index/Tables/DataList-874991.iwa",
    "Index/Tables/DataList-874992.iwa",
    "Index/Tables/DataList-874993.iwa",
    "Index/Tables/DataList-874994.iwa",
    "Index/Tables/DataList-875124.iwa",
    "Index/Tables/DataList-875132.iwa",
    "Index/Tables/HeaderStorageBucket-875122.iwa",
    "Index/Tables/HeaderStorageBucket-875123.iwa",
    "Index/Tables/HeaderStorageBucket-875130.iwa",
    "Index/Tables/HeaderStorageBucket-875131.iwa",
    "Index/Tables/Tile-874954.iwa",
    "Index/Tables/Tile-874983.iwa",
    "Index/ViewState.iwa",
    "Metadata/BuildVersionHistory.plist",
    "Metadata/DocumentIdentifier",
    "Metadata/Properties.plist",
    "preview-micro.jpg",
    "preview-web.jpg",
    "preview.jpg",
]
IMAGE_SIZES = [1512, 10069, 76608]


def test_read_simple():
    reader = file_utils.zip_file_reader(NUMBERS_FILE)
    sorted_filenames = sorted([filename for filename, handle in reader])
    assert sorted_filenames == FILE_LIST


def test_process_single_file():
    reader = file_utils.zip_file_reader(NUMBERS_FILE)

    results = {}

    def sink(filename, contents):
        results[filename] = contents

    for filename, handle in reader:
        if filename.endswith(".jpg"):
            file_utils.process_file(filename, handle, sink)

    assert len(results) == 3
    file_lengths = [len(c) for c in results.values()]
    assert file_lengths == IMAGE_SIZES
    is_jpeg = [c[6:10] == b"JFIF" for c in results.values()]
    assert is_jpeg == [True, True, True]
