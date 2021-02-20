mkfile_path := $(abspath $(lastword $(MAKEFILE_LIST)))
current_dir := $(notdir $(patsubst %/,%,$(dir $(mkfile_path))))

# Change this to the location of the proto-dump executable. Default assumes
# a repo in the same root as numbers-parser
PROTO_DUMP=$(current_dir})../proto-dump/build/Release/proto-dump

# Location of the Numbers application
NUMBERS=/Applications/Numbers.app

PROTO_SOURCES = $(wildcard protos/*.proto)
PROTO_CLASSES = $(patsubst protos/%.proto,numbers_parser/generated/%_pb2.py,$(PROTO_SOURCES))

.PHONY: all clean install test

all: $(PROTO_CLASSES) numbers_parser/generated/__init__.py

install: $(PROTO_CLASSES) numbers_parser/generated/__init__.py numbers_parser/*
	python3 setup.py install

upload: $(PROTO_CLASSES) numbers_parser/generated/__init__.py numbers_parser/*
	python3 setup.py upload

numbers_parser/generated:
	mkdir -p numbers_parser/generated
	# Note that if any of the incoming Protobuf definitions contain periods,
	# protoc will put them into their own Python packages. This is not desirable
	# for import rules in Python, so we replace non-final period characters with
	# underscores.
	python3 protos/rename_proto_files.py protos

numbers_parser/generated/%_pb2.py: protos/%.proto numbers_parser/generated
	protoc -I=protos --proto_path protos --python_out=numbers_parser/generated $<

numbers_parser/generated/__init__.py: numbers_parser/generated $(PROTO_CLASSES)
	touch $@
	# Huge hack for py3 support, see https://github.com/protocolbuffers/protobuf/issues/1491
	futurize --no-diffs --nobackups --both-stages --processes 4 -w numbers_parser/generated/

tmp/TSPRegistry.dump::
	rm -rf tmp
	protos/dump_mappings.sh "$(NUMBERS)"
	python3 protos/generate_mapping.py tmp/TSPRegistry.dump > numbers_parser/mapping.py

bootstrap: $(PROTO_DUMP) tmp/TSPRegistry.dump
	rm -f protos/*.proto
	PROTO_DUMP="$(PROTO_DUMP)" protos/extract_protos.sh "$(NUMBERS)"

clean:
	rm -rf numbers_parser/generated
	rm -rf numbers_parser.egg_info
	rm -rf dist

test: all
	python3 -m pytest . --cov=numbers_parser -W ignore::DeprecationWarning
