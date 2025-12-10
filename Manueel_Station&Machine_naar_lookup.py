import pyperclip

raw_input = """

B-4122140;#B-4122240	140CSH101;#240CSH101

"""

def clean_side(chunk: str) -> str:
    parts = [p.strip() for p in chunk.split(";#") if p.strip()]
    seen = set()
    cleaned = []
    for p in parts:
        if not p.startswith("B-"):
            p = "B-4122" + p
        if p not in seen:
            seen.add(p)
            cleaned.append(p)
    return "; ".join(cleaned)


output_lines = []

for line in raw_input.strip().splitlines():
    line = line.strip()
    if not line:
        continue

    if "\t" in line:
        left, right = line.split("\t", 1)
        left_clean = clean_side(left)
        right_clean = clean_side(right)
        output_lines.append(left_clean + "\t" + right_clean)
    else:
        output_lines.append(clean_side(line))

final = "\n".join(output_lines)

# copy straight into your clipboard
pyperclip.copy(final)

print("Copied!")


