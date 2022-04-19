#! /bin/bash

mkdir -p tmp
BINFILE="tmp/binfile"

me="$(cd "$(dirname "$0")" >/dev/null 2>&1; pwd -P)"
PREFIX=${PREFIX:-$(dirname "$me")}

function ok_echo {
  echo $(tput setaf 2)"$1"$(tput init)
}

function error_echo {
  echo $(tput setaf 1)"$1"$(tput init)
}

function fatal_error {
  error_echo "$1"
  exit 1
}

app="$1"
test ! -z "$app" || fatal_error "usage: $0 app-directory"
test "${app: -4}" == ".app" -a -d "$app/Contents" || fatal_error "$app: not an Application folder"

proto_dump_root="$(dirname "${PREFIX}")/proto-dump"
PROTO_DUMP=${PROTO_DUMP:-"$proto_dump_root/build/Release/proto-dump"}
test -x "$PROTO_DUMP" || fatal_error "$PROTO_DUMP: proto-dump found"

get_data() {
  while read -r data; do
	data=$(echo "$data" | sed 's/:.*Mach-O.*//')
	cat "$data" >> "$BINFILE"
  done
}

cd "$PREFIX"
rm -f "$BINFILE"

ok_echo "*** Searching $app/Contents for Mach-O binaries"
find "$app/Contents" -type f -print0 | xargs -0 file | grep 'Mach-O.*binary' | get_data

ok_echo "*** Dumping proto files to protos"
"$PROTO_DUMP" --output=tmp/protos "$BINFILE" || fatal_error "$PROTO_DUMP: error"

ok_echo "*** Renaming proto files"
python3 protos/rename_proto_files.py tmp/protos/binfile || fatal_error "rename_proto_files.py: error"

ok_echo "*** Cleaning up"
mv tmp/protos/binfile/*.proto protos
rm -rf tmp/binfile tmp/protos
