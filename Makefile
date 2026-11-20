.DELETE_ON_ERROR:

# --- Configuration ---

PACKAGE         := numbers-parser
package_c       := $(subst -,_,$(PACKAGE))

# Location of the protoc compiler executable
PROTOC          ?= /opt/homebrew/bin/protoc

# Location of the Numbers application
NUMBERS         ?= /Applications/Numbers Creator Studio.app

# Temporary re-signed copy of Numbers used during bootstrap
TMP_NUMBERS_APP ?= /tmp/Numbers.unsigned.app
ENTITLEMENTS    := src/build/entitlements.xml
TST_TABLES       = $(NUMBERS)/Contents/Frameworks/TSTables.framework/Versions/A/TSTables

# Deferred macOS-only lookups (only evaluated when bootstrap targets run)
IDENTITY         = $(shell security find-identity -v -p codesigning | head -n 1 | python3 -c 'import sys; print(sys.stdin.read().split("\"")[1])' 2>/dev/null)
LLDB_PYTHON_PATH = $(shell lldb --python-path)

# Terminal colors (safe in non-TTY / CI environments)
TTY_GREEN       := $(shell tput setaf 2 2>/dev/null || true)
TTY_RED         := $(shell tput setaf 1 2>/dev/null || true)
TTY_RESET       := $(shell tput sgr0 2>/dev/null || true)
info_message     = echo "$(TTY_GREEN)Bootstrap: $(1)$(TTY_RESET)"
error_message    = echo "$(TTY_RED)Bootstrap: $(1)$(TTY_RESET)"

DOCS_SOURCES = $(shell find docs -name '*.rst') \
               docs/conf.py \
               src/$(package_c)/*.py \
               docs/build/_static/custom.css

BOOTSTRAP_FILES = src/protos/TNArchives.proto \
                  src/$(package_c)/generated/TNArchives_pb2.py \
                  src/$(package_c)/generated/__init__.py \
                  src/$(package_c)/generated/mapping.py \
                  src/$(package_c)/generated/functionmap.py \
                  src/$(package_c)/generated/fontmap.py

GENERATED_PY_COPIES := mapping.py functionmap.py fontmap.py

.PHONY: all readme bootstrap clean veryclean test profile dist docs visualcheck check_prereqs check_identity

# --- Help ---

all:
	@echo "make targets:"
	@echo "    test        - run pytest with all tests"
	@echo "    readme      - rebuild README.md from source"
	@echo "    docs        - rebuild markdown and HTML docs from source"
	@echo "    profile     - run pytest and generate a profile graph"
	@echo "    dist        - build distributions"
	@echo "    clean       - delete temporary files for test, coverage, etc."
	@echo "    veryclean   - delete all auto-generated files (requires new bootstrap)"
	@echo "    bootstrap   - rebuild all auto-generated files for new Numbers version"
	@echo "    visualcheck - Generate a PDF from Numbers and check visual matches (macOS only)"

# --- Test & Quality ---

test:
	uv run --group dev pytest -n logical

profile:
	@set -e; \
	for sub_test in tables borders styles; do \
		test_name="test_profiling_$$sub_test"; \
		heat_graph="prof/$$test_name.svg"; \
		echo "======== Profile run: $$test_name"; \
		uv run --group profile pytest --profile-svg -q --experimental --no-cov tests/test_large.py -k "$$test_name"; \
		mv prof/combined.svg "$$heat_graph"; \
		echo ">>>>>>>> Saved to $$heat_graph"; \
	done

visualcheck:
	uv run --group visualcheck python3 src/build/check_visual_document.py

# --- Docs & Distribution ---

dist: readme
	rm -rf dist
	uv build

docs: docs/build/index.html docs/build/index.md
	uv run --group docs python3 src/build/check_docs_protobufs.py --doc tests/data/check-docs-protobufs.numbers
	uv run --group docs python3 src/build/check_docs_links.py docs/build
	uv run --group docs python3 src/build/check_docs_examples.py

docs/build/_static/custom.css: docs/custom.css
	@mkdir -p $(dir $@)
	@cp "$<" "$@"

docs/build/index.html: docs/build/index.md $(DOCS_SOURCES)
	@mkdir -p $(dir $@)
	uv run --group docs sphinx-build -q -b html -t HtmlDocs docs docs/build

docs/build/index.md: $(DOCS_SOURCES)
	@mkdir -p $(dir $@)
	uv run --group docs sphinx-build -q -b markdown -t MarkdownDocs docs docs/build docs/index.rst

readme: README.md

README.md: docs/build/index.md
	@cp "$<" "$@"

# --- Bootstrap ---

bootstrap: check_prereqs check_identity $(BOOTSTRAP_FILES)

check_identity:
	@if [ -z "$(IDENTITY)" ]; then \
		$(call error_message,"failure: no code signing identity found"); \
		exit 1; \
	fi

check_prereqs:
	@if ! command -v codesign >/dev/null 2>&1 || ! codesign -d "$$(command -v codesign)" >/dev/null 2>&1; then \
		$(call error_message,"failure: codesign is not on PATH or not working"); \
		exit 1; \
	fi
	@if ! "$(PROTOC)" --version >/dev/null 2>&1; then \
		$(call error_message,"failure: protoc not found or not working at $(PROTOC)"); \
		exit 1; \
	fi
	@if [ ! -d "$(NUMBERS)" ]; then \
		$(call error_message,"failure: Numbers.app not found at $(NUMBERS)"); \
		exit 1; \
	fi

$(TMP_NUMBERS_APP): $(ENTITLEMENTS)
	@$(call info_message,"creating local copy of Numbers.app")
	@rm -rf "$@"
	@cp -R "$(NUMBERS)" "$@"
	@$(call info_message,"re-signing local Numbers.app")
	@xattr -cr "$@"
	@codesign --remove-signature "$@"
	@codesign --entitlements "$(ENTITLEMENTS)" --sign "$(IDENTITY)" "$@"
	@$(call info_message,"checking signatures of local Numbers.app")
	@codesign --verify "$@"

.bootstrap/mapping.json: $(TMP_NUMBERS_APP)
	@$(call info_message,"dumping protobuf IDs using lldb")
	@mkdir -p $(dir $@)
	@PYTHONPATH="$(LLDB_PYTHON_PATH):src" xcrun python3 \
		src/build/extract_mapping.py \
		"$(TMP_NUMBERS_APP)/Contents/MacOS/Numbers" "$@"

.bootstrap/mapping.py: .bootstrap/mapping.json
	@$(call info_message,"creating Python mapping of Protobuf IDs")
	@mkdir -p $(dir $@)
	uv run --group bootstrap python3 src/build/generate_mapping.py "$<" "$@"

.bootstrap/functionmap.py:
	@$(call info_message,"extracting function names from Numbers")
	@mkdir -p $(dir $@)
	uv run --group bootstrap python3 src/build/extract_functions.py "$(TST_TABLES)" "$@" >/dev/null

.bootstrap/fontmap.py:
	@$(call info_message,"generating font name map")
	@mkdir -p $(dir $@)
	uv run --group bootstrap python3 src/build/generate_fontmap.py "$@"

.bootstrap/protos/TNArchives.proto:
	@$(call info_message,"extracting protobufs from Numbers")
	@mkdir -p $(dir $@)
	uv run --group bootstrap python3 src/build/protodump.py "$(NUMBERS)" .bootstrap/protos
	uv run --group bootstrap python3 src/build/rename_proto_files.py .bootstrap/protos

$(addprefix src/$(package_c)/generated/,$(GENERATED_PY_COPIES)): src/$(package_c)/generated/%.py: .bootstrap/%.py
	@mkdir -p $(dir $@)
	@cp "$<" "$@"

src/$(package_c)/generated/TNArchives_pb2.py: .bootstrap/protos/TNArchives.proto
	@$(call info_message,"compiling Python packages from protobufs")
	@mkdir -p $(dir $@)
	@set -e; \
	for proto in .bootstrap/protos/*.proto; do \
		uv run --group bootstrap "$(PROTOC)" -I=.bootstrap/protos \
			--python_out=src/$(package_c)/generated \
			--mypy_out=src/$(package_c)/generated "$$proto"; \
	done

src/protos/TNArchives.proto: .bootstrap/protos/TNArchives.proto
	@$(call info_message,"creating git-tracked copies of protos")
	@mkdir -p src/protos
	@set -e; \
	for proto in .bootstrap/protos/*.proto; do \
		cp "$$proto" src/protos/; \
	done

src/$(package_c)/generated/__init__.py: src/$(package_c)/generated/TNArchives_pb2.py
	@$(call info_message,"patching paths in generated protobuf files")
	uv run --group bootstrap python3 src/build/replace_paths.py src/$(package_c)/generated/T*.py
	@touch "$@"

# --- Cleanup ---

veryclean: clean
	rm -rf .bootstrap "$(TMP_NUMBERS_APP)"
	rm -f src/$(package_c)/generated/__init__.py \
	      src/$(package_c)/generated/mapping.py \
	      src/$(package_c)/generated/*_pb2.py \
	      src/$(package_c)/generated/*_pb2.pyi

clean:
	rm -rf src/$(package_c).egg-info
	rm -rf coverage_html_report coverage.xml
	rm -rf dist
	rm -rf docs/build
	rm -rf .tox
	rm -rf .pytest_cache
	rm -f *.txt *.numbers *.png *.log
