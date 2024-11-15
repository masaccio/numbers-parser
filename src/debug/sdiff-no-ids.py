import re
import subprocess
from os import get_terminal_size
from sys import argv, exit, stderr

from colorama import Fore, Style

diff_version = subprocess.check_output(["diff", "--version"])  # noqa: S603, S607
if "GNU diffutils" not in diff_version.decode():
    print("Unsupported diff utility (expected GNU diffutils)", file=stderr)
    exit(1)

try:
    diff_width = get_terminal_size().columns - 5
except OSError:
    diff_width = 160

args = ["diff", "--side-by-side", f"--width={diff_width}"] + argv[1:]
proc = subprocess.run(args, capture_output=True, encoding="utf8", check=False)  # noqa: S603
if proc.stderr:
    print(proc.stderr, file=stderr)
    exit(1)

for line in proc.stdout.splitlines():
    if m := re.search(r"(.*)(\s\|\s)(.*)", line):
        left_side = m.group(1)
        space = m.group(2)
        right_side = m.group(3)

        left_side_no_numbers = re.sub(r"\b(\d+)\b", lambda x: "X" * len(x.group(0)), left_side)
        right_side_no_numbers = re.sub(r"\b(\d+)\b", lambda x: "X" * len(x.group(0)), right_side)
        if left_side_no_numbers[0 : len(right_side_no_numbers)] == right_side_no_numbers:
            left_side = re.sub(
                r"\b(\d+)\b",
                lambda x: Fore.GREEN + x.group(0) + Fore.RESET,
                left_side,
            )
            right_side = re.sub(
                r"\b(\d+)\b",
                lambda x: Fore.GREEN + x.group(0) + Fore.RESET,
                right_side,
            )
            print(left_side, space.replace("|", "~"), right_side)
            continue

        color_on = False
        left_side_colored = ""
        right_side_colored = ""
        for i in range(len(left_side)):
            if i >= len(right_side):
                left_side_colored += left_side[i:]
                break
            if left_side[i] == right_side[i] and color_on:
                left_side_colored += Style.RESET_ALL
                right_side_colored += Style.RESET_ALL
                color_on = False
            elif left_side[i] != right_side[i] and not color_on:
                color_on = True
                left_side_colored += Fore.RED + Style.BRIGHT
                right_side_colored += Fore.RED + Style.BRIGHT

            left_side_colored += left_side[i]
            right_side_colored += right_side[i]

        left_side_colored += Style.RESET_ALL
        right_side_colored += Style.RESET_ALL
        print(left_side_colored + space + right_side_colored)
    else:
        print(line)
