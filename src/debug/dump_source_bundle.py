import os

for root_path, bundle_name in [
    ("src/numbers_parser", "source"),
    ("docs", "docs"),
    ("tests", "tests"),
]:
    out_filename = f"{bundle_name}_review_bundle.txt"
    with open(out_filename, "w", encoding="utf-8") as out:
        for root, _, files in os.walk(root_path):
            if "generated" in root:
                continue
            for f in files:
                if f.endswith((".py", ".rst")):
                    path = os.path.join(root, f)
                    out.write(f"\n\n=== FILE: {path} ===\n\n")
                    with open(path, encoding="utf-8", errors="ignore") as fh:
                        out.write(fh.read())
    print(f"Created {out_filename}")
