#! /usr/bin/env python3

from binascii import hexlify
from struct import pack
from sys import argv

for arg in argv[1:]:
  print(arg, "=", hexlify(pack("<d", float(arg)), sep=":"))
