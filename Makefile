mkfile_path := $(abspath $(lastword $(MAKEFILE_LIST)))
current_dir := $(notdir $(patsubst %/,%,$(dir $(mkfile_path))))

# Change this to the location of the proto-dump executable. Default assumes
# a repo in the same root as numbers-parser
PROTO_DUMP=$(current_dir})../proto-dump/build/Release/proto-dump

# Location of the Numbers application
NUMBERS=/Applications/Numbers.app

PROTO_SOURCES = $(wildcard protos/*.proto)
PROTO_CLASSES = $(patsubst protos/%.proto,src/numbers_parser/generated/%_pb2.py,$(PROTO_SOURCES))

SOURCE_FILES=src/numbers_parser/generated/__init__.py $(wildcard src/numbers_parser/*.py)
RELEASE_TARBALL=dist/numbers-parser-$(shell python3 setup.py --version).tar.gz

.PHONY: all clean install test

all: $(PROTO_CLASSES) src/numbers_parser/generated/__init__.py

install: $(PROTO_CLASSES) $(SOURCE_FILES)
	python3 setup.py install

$(info SOURCE_FILES=$(SOURCE_FILES))
$(RELEASE_TARBALL): $(SOURCE_FILES)
	python3 setup.py sdist
	tox

upload: $(RELEASE_TARBALL)
	twine upload $(RELEASE_TARBALL)

src/numbers_parser/generated:
	mkdir -p src/numbers_parser/generated
	# Note that if any of the incoming Protobuf definitions contain periods,
	# protoc will put them into their own Python packages. This is not desirable
	# for import rules in Python, so we replace non-final period characters with
	# underscores.
	python3 protos/rename_proto_files.py protos

src/numbers_parser/generated/%_pb2.py: protos/%.proto src/numbers_parser/generated
	protoc -I=protos --proto_path protos --python_out=src/numbers_parser/generated $<

src/numbers_parser/generated/__init__.py: src/numbers_parser/generated $(PROTO_CLASSES)
	touch $@
	python3 protos/replace_paths.py src/numbers_parser/generated/T*.py

tmp/TSPRegistry.dump::
	rm -rf tmp
	protos/dump_mappings.sh "$(NUMBERS)"

src/numbers_parser/mapping.py: $(PROTO_CLASSES)
	python3 protos/generate_mapping.py tmp/TSPRegistry.dump > src/numbers_parser/mapping.py

bootstrap: $(PROTO_DUMP) tmp/TSPRegistry.dump
	rm -f protos/*.proto
	rm -f src/numbers_parser/mapping.py
	PROTO_DUMP="$(PROTO_DUMP)" protos/extract_protos.sh "$(NUMBERS)"
	$(MAKE) all src/numbers_parser/mapping.py
	rm -rf tmp

clean:
	rm -rf src/numbers_parser/generated
	rm -rf numbers_parser.egg-info
	rm -rf coverage_html_report
	rm -rf dist

test: all
	python3 -m pytest . --cov=numbers_parser -W ignore::DeprecationWarning
