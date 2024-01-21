# Change this to the name of a code-signing certificate. A self-signed
# certificate is suitable for this.
IDENTITY=Jonathan Connell
#
# Change this to the location of the proto-dump executable
PROTOC=/usr/local/bin/protoc
#
# Location of the Numbers application
NUMBERS=/Applications/Numbers.app

# Xcode version of Python that includes LLDB package
LLDB_PYTHON_PATH := ${shell lldb --python-path}

PACKAGE=numbers-parser
package_c := $(subst -,_,$(PACKAGE))

.PHONY: clean veryclean test coverage profile sdist upload docs

all:
	@echo "make targets:"
	@echo "    test       - run pytest with all tests"
	@echo "    docs       - rebuild maekdown docs from source"
	@echo "    profile    - run pytest and generate a profile graph"
	@echo "    dist       - build distributions"
	@echo "    upload     - upload package to PyPI"
	@echo "    clean      - delete temporary files for test, coverage, etc."
	@echo "    veryclean  - delete all auto-generated files (requires new bootstrap)"
	@echo "    bootstrap  - rebuild all auto-generated files for new Numbers version"

dist:
	poetry build

upload:
	tox
	poetry publish --build

profile:
	poetry run pytest --profile
	poetry run gprof2dot -f pstats prof/combined.prof | dot -Tpng -o prof/combined.png

docs/build/index.html: docs/index.rst docs/conf.py src/$(package_c)/*.py
	@mkdir -p docs/build
	poetry run sphinx-build -q -b html  docs docs/build


README.md: docs/build/index.md
	cp $< $@

test:
	poetry run pytest -n logical

BOOTSTRAP_FILES = src/$(package_c)/generated/functionmap.py \
				  src/$(package_c)/generated/fontmap.py \
				  src/$(package_c)/generated/__init__.py \
				  src/$(package_c)/mapping.py \
				  src/protos/TNArchives.proto

bootstrap: $(BOOTSTRAP_FILES)

ENTITLEMENTS = src/build/entitlements.xml

.bootstrap/Numbers.unsigned.app: $(ENTITLEMENTS)
	@echo $$(tput setaf 2)"Bootstrap: extracting protobuf mapping from Numbers"$$(tput init)
	@mkdir -p .bootstrap
	rm -rf $@
	cp -R $(NUMBERS) $@
	xattr -cr $@
	codesign --remove-signature $@
	codesign --entitlements $(ENTITLEMENTS) --sign "${IDENTITY}" $@
	codesign --verify $@

.bootstrap/mapping.json: .bootstrap/Numbers.unsigned.app
	@mkdir -p .bootstrap
	PYTHONPATH=${LLDB_PYTHON_PATH}:src xcrun python3 \
		src/build/extract_mapping.py \
		.bootstrap/Numbers.unsigned.app/Contents/MacOS/Numbers $@

.bootstrap/mapping.py: .bootstrap/mapping.json
	@mkdir -p $(dir $@)
	python3 src/build/generate_mapping.py $< $@

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
	poetry run python3 src/build/extract_functions.py $(TST_TABLES) $@

.bootstrap/fontmap.py:
	@echo $$(tput setaf 2)"Bootstrap: generating font name map"$$(tput init)
	@mkdir -p .bootstrap
	poetry run python3 src/build/generate_fontmap.py $@

.bootstrap/protos/TNArchives.proto:
	@echo $$(tput setaf 2)"Bootstrap: extracting protobufs from Numbers"$$(tput init)
	PROTOCOL_BUFFERS_PYTHON_IMPLEMENTATION=cpp python3 src/build/protodump.py $(NUMBERS) .bootstrap/protos
	poetry run python3 src/build/rename_proto_files.py .bootstrap/protos

src/$(package_c)/mapping.py: .bootstrap/mapping.py
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
	python3 src/build/replace_paths.py src/$(package_c)/generated/T*.py
	touch $@

veryclean:
	make clean
	rm -rf .bootstrap
	rm -rf src/$(package_c)/generated

clean:
	rm -rf src/$(package_c).egg-info
	rm -rf coverage_html_report
	rm -rf dist
	rm -rf docs/build
	rm -rf .tox
	rm -rf .pytest_cache
