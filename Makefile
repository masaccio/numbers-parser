PROTO_SOURCES = $(wildcard protos/*.proto)
PROTO_CLASSES = $(patsubst protos/%.proto,keynote_parser/generated/%_pb2.py,$(PROTO_SOURCES))

.PHONY: all clean install

all: $(PROTO_CLASSES) keynote_parser/generated/__init__.py

install: $(PROTO_CLASSES) keynote_parser/generated/__init__.py keynote_parser/*
	python setup.py install

upload: $(PROTO_CLASSES) keynote_parser/generated/__init__.py keynote_parser/*
	python setup.py upload

keynote_parser/generated:
	mkdir -p keynote_parser/generated

keynote_parser/generated/%_pb2.py: protos/%.proto keynote_parser/generated
	protoc -I=protos --python_out=keynote_parser/generated $<

keynote_parser/generated/__init__.py: keynote_parser/generated
	touch $@

clean:
	rm -rf keynote_parser/generated
	rm -rf keynote_parser.egg_info
	rm -rf dist
