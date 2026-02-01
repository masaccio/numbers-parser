# Apple developer certificate for re-signing
IDENTITY := $(shell security find-identity -v -p codesigning | head -n 1 | python3 -c 'import sys; print(sys.stdin.read().split("\"")[1])' 2>/dev/null)

# Change this to the location of the proto-dump executable
PROTOC=/opt/homebrew/bin/protoc

# Location of the Numbers application
NUMBERS=/Applications/Numbers.app

# Xcode version of Python that includes LLDB package
LLDB_PYTHON_PATH := ${shell lldb --python-path}

PACKAGE=numbers-parser
package_c := $(subst -,_,$(PACKAGE))

.PHONY: clean veryclean test coverage profile dist upload docs

all:
	@echo "make targets:"
	@echo "    test       - run pytest with all tests"
	@echo "    readme     - rebuild README.md from source"
	@echo "    docs       - rebuild markdown and HTML docs from source"
	@echo "    profile    - run pytest and generate a profile graph"
	@echo "    dist       - build distributions"
	@echo "    upload     - upload package to PyPI"
	@echo "    clean      - delete temporary files for test, coverage, etc."
	@echo "    veryclean  - delete all auto-generated files (requires new bootstrap)"
	@echo "    bootstrap  - rebuild all auto-generated files for new Numbers version"

dist: readme
	-rm -rf dist
	uv build

upload:
	-rm -rf dist
	tox
	uv build
	uv publish

DOCS_SOURCES = $(shell find docs -name \*.rst) \
			   docs/conf.py \
			   src/$(package_c)/*.py \
			   docs/build/_static/custom.css

docs: docs/build/index.html

docs/build/_static/custom.css: docs/custom.css
	mkdir -p docs/build/_static
	cp $< $@

docs/build/index.html: $(DOCS_SOURCES)
	@mkdir -p docs/build
	uv sync --group docs
	uv run sphinx-build -q -b html -t HtmlDocs docs docs/build
	uv run sphinx-build -q -b markdown -t MarkdownDocs docs docs/build docs/index.rst

readme:
	@mkdir -p docs/build
	uv sync --group docs
	uv run sphinx-build -q -b markdown -t MarkdownDocs docs docs/build docs/index.rst
	cp docs/build/index.md README.md

profile:
	uv run pytest --profile
	uv run gprof2dot -f pstats prof/combined.prof | dot -Tpng -o prof/combined.png


test:
	uv run pytest -n logical

BOOTSTRAP_FILES = src/$(package_c)/generated/functionmap.py \
				  src/$(package_c)/generated/fontmap.py \
				  src/$(package_c)/generated/__init__.py \
				  src/$(package_c)/generated/mapping.py \
				  src/protos/TNArchives.proto

bootstrap: check_identity $(BOOTSTRAP_FILES)

check_identity::
	@if [ "$(IDENTITY)" == "" ]; then \
		echo $$(tput setaf 1)"Bootstrap: No code signing identity found."$$(tput init); \
		exit 1; \
	fi

ENTITLEMENTS = src/build/entitlements.xml
TMP_NUMBERS_APP = /tmp/Numbers.unsigned.app

$(TMP_NUMBERS_APP): $(ENTITLEMENTS)
	@echo $$(tput setaf 2)"Bootstrap: extracting protobuf mapping from Numbers"$$(tput init)
	@mkdir -p .bootstrap
	rm -rf $@
	cp -R $(NUMBERS) $@
	xattr -cr $@
	codesign --remove-signature $@
	codesign --entitlements $(ENTITLEMENTS) --sign "${IDENTITY}" $@
	codesign --verify $@

.bootstrap/mapping.json: $(TMP_NUMBERS_APP)
	@mkdir -p .bootstrap
	PYTHONPATH=${LLDB_PYTHON_PATH}:src xcrun python3 \
		src/build/extract_mapping.py \
		$(TMP_NUMBERS_APP)/Contents/MacOS/Numbers $@

.bootstrap/mapping.py: .bootstrap/mapping.json
	@mkdir -p $(dir $@)
	uv run python3 src/build/generate_mapping.py $< $@

src/$(package_c)/generated/functionmap.py: .bootstrap/functionmap.py
	@mkdir -p $(dir $@)
	cp $< $@

src/$(package_c)/generated/fontmap.py: .bootstrap/fontmap.py
	@mkdir -p $(dir $@)
	cp $< $@

TST_TABLES=$(NUMBERS)/Contents/Frameworks/TSTables.framework/Versions/A/TSTables
.bootstrap/functionmap.py:
	@echo $$(tput setaf 2)"Bootstrap: extracting function names from Numbers"$$(tput init)
	@mkdir -p .bootstrap
	uv run python3 src/build/extract_functions.py $(TST_TABLES) $@ >/dev/null

.bootstrap/fontmap.py:
	@echo $$(tput setaf 2)"Bootstrap: generating font name map"$$(tput init)
	@mkdir -p .bootstrap
	uv sync --group bootstrap
	uv run python3 src/build/generate_fontmap.py $@

.bootstrap/protos/TNArchives.proto:
	@echo $$(tput setaf 2)"Bootstrap: extracting protobufs from Numbers"$$(tput init)
	uv run python3 src/build/protodump.py $(NUMBERS) .bootstrap/protos
	uv run python3 src/build/rename_proto_files.py .bootstrap/protos

src/$(package_c)/generated/mapping.py: .bootstrap/mapping.py
	cp $< $@

src/$(package_c)/generated/TNArchives_pb2.py: .bootstrap/protos/TNArchives.proto
	@echo $$(tput setaf 2)"Bootstrap: compiling Python packages from protobufs"$$(tput init)
	@mkdir -p src/$(package_c)/generated
	for proto in .bootstrap/protos/*.proto; do \
	    $(PROTOC) -I=.bootstrap/protos --proto_path .bootstrap/protos --python_out=src/$(package_c)/generated $$proto; \
	done

src/protos/TNArchives.proto: .bootstrap/protos/TNArchives.proto
	@echo $$(tput setaf 2)"Bootstrap: creating git-tracked copies of protos"$$(tput init)
	@mkdir -p src/protos
	for proto in .bootstrap/protos/*.proto; do \
		cp $$proto src/protos; \
	done


src/$(package_c)/generated/__init__.py: src/$(package_c)/generated/TNArchives_pb2.py
	@echo $$(tput setaf 2)"Bootstrap: patching paths in generated protobuf files"$$(tput init)
	uv run python3 src/build/replace_paths.py src/$(package_c)/generated/T*.py
	touch $@

veryclean:
	make clean
	rm -rf .bootstrap
	rm -rf src/$(package_c)/generated
	rm -rf $(TMP_NUMBERS_APP)

clean:
	rm -rf src/$(package_c).egg-info
	rm -rf coverage_html_report
	rm -rf dist
	rm -rf docs/build
	rm -rf .tox
	rm -rf .pytest_cache
