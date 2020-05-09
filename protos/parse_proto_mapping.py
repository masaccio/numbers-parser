"""
Passed a filename that contains lines output by [TSPRegistry sharedRegistry] in lldb/Xcode:

 	182 -> 0x10f536100 KN.CommandSlideResetMasterBackgroundObjectsArchive
  	181 -> 0x10f5362e0 KN.ActionGhostSelectionTransformerArchive
  	...

...this script will print a sorted JSON object definition of that mapping from class
definition ID to Protobuf message type name. This is some of my hackiest code. Please
don't use this for anything important. It's mostly here for next time I need it.
"""
import sys
import json


def parse_proto_mapping(input_file):
    split = [x.strip().split(" -> ") for x in open(input_file).split("\n")]
    print(
        json.dumps(
            dict(sorted([(int(a), b.split(" ")[-1]) for a, b in split if 'null' not in b])),
            indent=2,
        )
    )


if __name__ == "__main__":
    parse_proto_mapping(sys.argv[-1])
