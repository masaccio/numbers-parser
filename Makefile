mkfile_path := $(abspath $(lastword $(MAKEFILE_LIST)))
current_dir := $(notdir $(patsubst %/,%,$(dir $(mkfile_path))))

# Change this to the location of the proto-dump executable. Default assumes
# a repo in the same root as numbers-parser
PROTO_DUMP=$(current_dir})../proto-dump/build/Release/proto-dump

# Location of the Numbers application
NUMBERS=/Applications/Numbers.app

PROTO_SOURCES = $(wildcard protos/*.proto)
PROTO_CLASSES = $(patsubst protos/%.proto,src/numbers_parser/generated/%_pb2.py,$(PROTO_SOURCES))

SOURCE_FILES=src/numbers_parser/generated/__init__.py $(wildcard src/numbers_parser/*.py) $(wildcard scripts/*)
RELEASE_TARBALL=dist/numbers-parser-$(shell python3 setup.py --version).tar.gz

.PHONY: all clean veryclean install test coverage sdist upload

all: $(PROTO_CLASSES) src/numbers_parser/generated/__init__.py

install: $(SOURCE_FILES)
	python3 setup.py install

$(RELEASE_TARBALL):
	python3 setup.py sdist

upload: $(RELEASE_TARBALL)
	tox
	twine upload $(RELEASE_TARBALL)

test: all
	PYTHONPATH=src python3 -m pytest tests

coverage: all
	PYTHONPATH=src python3 -m pytest --cov=numbers_parser --cov-report=html

src/numbers_parser/generated/%_pb2.py: protos/%.proto 
	@mkdir -p src/numbers_parser/generated
	protoc -I=protos --proto_path protos --python_out=src/numbers_parser/generated $<

src/numbers_parser/generated/__init__.py: $(PROTO_SOURCES)
	touch $@
	python3 protos/replace_paths.py src/numbers_parser/generated/T*.py

protos/TSPRegistry.dump:
	protos/dump_mappings.sh "$(NUMBERS)"

src/numbers_parser/mapping.py: $(PROTO_CLASSES)
	python3 protos/generate_mapping.py protos/TSPRegistry.dump > src/numbers_parser/mapping.py

bootstrap: $(PROTO_DUMP) protos/TSPRegistry.dump
	rm -f protos/*.proto
	rm -f src/numbers_parser/mapping.py
	PROTO_DUMP="$(PROTO_DUMP)" protos/extract_protos.sh "$(NUMBERS)"
	python3 protos/rename_proto_files.py protos
	$(MAKE) all src/numbers_parser/mapping.py
	rm -rf tmp

# Deleting TSPRegistry.dump will require System Integrity Protection to
# be disabled to then recreate using lldb
veryclean:
	$(MAKE) clean
	rm -f protos/*.protos
	rm -f protos/TSPRegistry.dump

clean:
	rm -rf src/numbers_parser/generated
	rm -rf numbers_parser.egg-info
	rm -rf coverage_html_report
	rm -rf dist
	rm -rf build
	rm -rf .tox
