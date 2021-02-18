#! /bin/bash

app="$1"

if [ -z "$app" ]; then
  echo "usage: $0 app-directory"
  exit 1
fi

if [ "${app: -4}" != ".app" -o ! -d "$app/Contents" ]; then
  echo "$app: not an Application folder"
  exit 1
fi

get_data() {
  binfile="$1"
  while read -r data; do
	data=$(echo "$data" | sed 's/:.*Mach-O.*//')
	echo "Copying $data to $binfile"
	cat "$data" >> "$binfile"
  done
}

rm -rf binfile protos
mkdir protos

echo "*** Searching $app/Contents for Mach-O binaries"
find "$app/Contents" -type f -print0 | xargs -0 file | grep 'Mach-O.*binary' | get_data "binfile"

echo "*** Dumping proto files to protos"
../../proto-dump/build/Release/proto-dump --output=protos binfile

echo "*** Renaming proto files"
python3 rename_proto_files.py protos/binfile

echo "*** Cleaning up"
mv protos/binfile/*.proto .
rm -rf protos binfile
