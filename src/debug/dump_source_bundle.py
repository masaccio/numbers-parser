import os

with open("review_bundle.txt", "w", encoding="utf-8") as out:
    for root, _, files in os.walk("."):
        for f in files:
            if f.endswith((".py", ".rst")) and root in [
                "./src/numbers_parser",
                "./docs",
                "./docs/api",
            ]:
                print(root, f)
                path = os.path.join(root, f)
                out.write(f"\n\n=== FILE: {path} ===\n\n")
                with open(path, encoding="utf-8", errors="ignore") as fh:
                    out.write(fh.read())
print("Created review_bundle.txt")
