#!/usr/bin/xcrun python3
"""
lldb script to dump TSPersistence registry mapping from iWork apps.

Licensed under the MIT license.

Copyright 2020-2022 Peter Sobot
Copyright 2022 Jon Connell
Copyright 2022 SheetJS LLC
"""

import json
import os
import sys

import lldb
from debug.lldbutil import print_stacktrace

if len(sys.argv) != 3:
    raise (ValueError(f"Usage: {sys.argv[0]} exe-file output.json"))

exe = sys.argv[1]
output = sys.argv[2]

debugger = lldb.SBDebugger.Create()
debugger.SetAsync(False)
target = debugger.CreateTargetWithFileAndArch(exe, None)

# Note: original script also created breakpoints on _handleAEOpenEvent
# but that is too early in Numbers 12.1
target.BreakpointCreateByName("-[NSApplication _sendFinishLaunchingNotification]")
target.BreakpointCreateByName("-[NSApplication _crashOnException:]")

# Note: original script skipped [CKContainer containerWithIdentifier:]
target.BreakpointCreateByRegex("CloudKit")

process = target.LaunchSimple(None, None, os.getcwd())
if not process:
    raise ValueError("Failed to launch process: " + exe)

if process.GetState() == lldb.eStateExited:
    raise ValueError(f"LLDB was unable to stop process! {process}")

try:
    while process.GetState() == lldb.eStateStopped:
        thread = process.GetThreadAtIndex(0)
        frame = thread.GetSelectedFrame()
        if frame.name == "-[NSApplication _crashOnException:]":
            raise ValueError(f"Process crashed at {frame.name}")

        stop_reason = thread.GetStopReason()

        if stop_reason == lldb.eStopReasonException:
            print_stacktrace(thread)
            function = frame.GetFunction()
            function_or_symbol = function if function else frame.GetSymbol()
            raise ValueError(f"Exception at {frame.name}")
        if stop_reason != lldb.eStopReasonBreakpoint:
            process.Continue()
            continue
        if frame.name[-8:] == "CloudKit":
            thread.ReturnFromFrame(
                thread.GetSelectedFrame(),
                lldb.SBValue().CreateValueFromExpression("0", ""),
            )
            process.Continue()
        elif frame.name == "-[NSApplication _sendFinishLaunchingNotification]":
            registry = frame.EvaluateExpression("[TSPRegistry sharedRegistry]")
            error = registry.GetError()
            if error.fail or registry.description is None:
                continue
                # raise (ValueError("Failed to extract registry"))
            split = [
                x.strip().split(" -> ")
                for x in registry.description.split("{")[1].split("}")[0].split("\n")
                if x.strip()
            ]
            json_str = json.dumps(
                dict(sorted([(int(a), b.split(" ")[-1]) for a, b in split if "null" not in b])),
                indent=2,
            )
            with open(output, "w") as fh:
                fh.write(json_str)
            break
        else:
            process.Continue()
finally:
    process.Kill()
