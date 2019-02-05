from keynote_parser import codec, file_utils

SIMPLE_FILENAME = './tests/data/simple-oneslide.key'
PHOTO_FILENAME = './tests/data/photo-oneslide.key'

SIMPLE_SLIDE_KEY = 'Index/Slide-8060.iwa'
SIMPLE_CONTENTS = [
    'Data/110809_familychineseoahu_en_00317_2040x1360-small-13.jpeg',
    'Data/110809_familychineseoahu_en_02016_981x654-small-11.jpeg',
    'Data/110809_familychineseoahu_en_02390_2880x1921-small-9.jpeg',
    'Data/mt-07EEAA23-F087-48C1-B6D0-04F3123EF6C7-222.jpg',
    'Data/mt-28415F3F-1E88-49CA-9288-73C5EC0E73AB-225.jpg',
    'Data/mt-28D0C904-1282-4458-A229-93C6451C7CE8-223.jpg',
    'Data/mt-2B1EC37A-AC99-407C-80A1-BFA6DAD07337-226.jpg',
    'Data/mt-6574C108-BEFF-4758-89D5-39BDC991689B-217.jpg',
    'Data/mt-948A42DF-8C2D-41B9-8BF0-4A87F169BBA1-221.jpg',
    'Data/mt-A672DE80-9545-4D04-8F0A-810E1439CC90-218.jpg',
    'Data/mt-CC8F34D9-65B4-4E66-85DF-0F5DC90B1F33-220.jpg',
    'Data/mt-E2298BC6-2981-4132-B2DD-DF3EFA1F2546-227.jpg',
    'Data/mt-E8BB6D9E-506A-42C8-A335-455FEDC2E11C-219.jpg',
    'Data/mt-F39BAD95-4ABB-4F39-A54F-1C6A46BCA6C5-224.jpg',
    'Data/st-7D643FA7-9A1F-45A5-A30F-7828735F3C35-205.jpg',
    'Data/st-7D643FA7-9A1F-45A5-A30F-7828735F3C35-237.jpg',
    'Index/AnnotationAuthorStorage.iwa',
    'Index/CalculationEngine.iwa',
    'Index/Document.iwa',
    'Index/DocumentMetadata.iwa',
    'Index/DocumentStylesheet.iwa',
    'Index/MasterSlide-7880.iwa',
    'Index/MasterSlide-7900.iwa',
    'Index/MasterSlide-7915.iwa',
    'Index/MasterSlide-7928.iwa',
    'Index/MasterSlide-7942.iwa',
    'Index/MasterSlide-7956.iwa',
    'Index/MasterSlide-7968.iwa',
    'Index/MasterSlide-7985.iwa',
    'Index/MasterSlide-7997.iwa',
    'Index/MasterSlide-8015.iwa',
    'Index/MasterSlide-8034.iwa',
    'Index/MasterSlide-8048.iwa',
    'Index/Metadata.iwa',
    'Index/Slide-8060.iwa',
    'Index/ViewState.iwa',
    'Metadata/BuildVersionHistory.plist',
    'Metadata/DocumentIdentifier',
    'Metadata/Properties.plist',
    'preview-micro.jpg',
    'preview-web.jpg',
    'preview.jpg']


def test_read_simple():
    reader = file_utils.zip_file_reader(SIMPLE_FILENAME, progress=False)
    sorted_files = sorted([filename for filename, handle in reader])
    assert sorted_files == SIMPLE_CONTENTS


def test_process_single_file():
    reader = file_utils.zip_file_reader(SIMPLE_FILENAME, progress=False)

    results = {}

    def sink(filename, contents):
        results[filename] = contents

    for filename, handle in reader:
        if filename == SIMPLE_SLIDE_KEY:
            file_utils.process_file(filename, handle, sink)

    assert len(results) == 1
    assert SIMPLE_SLIDE_KEY in results
    assert isinstance(results[SIMPLE_SLIDE_KEY], codec.IWAFile)

    chunks = results[SIMPLE_SLIDE_KEY].chunks
    assert len(chunks) == 1

    archives = results[SIMPLE_SLIDE_KEY].chunks[0].archives
    assert len(archives) == 11

    archive = archives[0]
    assert archive.header.identifier == 8060
