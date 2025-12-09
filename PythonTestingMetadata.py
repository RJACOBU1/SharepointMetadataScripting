#PythonTestingMetadata

raw_input = """
110CC102;#120CC106;#130CC101;#220CC106;#230CC101;#250CC102;#130CC102
110CC102;#120CC106;#130CC101;#220CC106;#230CC101;#250CC102;#130CC102
"""  # paste your full blob here

output_lines = []

for line in raw_input.strip().splitlines():
    parts = [p.strip() for p in line.split(";#") if p.strip()]
    # dedupe while preserving order
    seen = set()
    cleaned = []
    for p in parts:
        if not p.startswith("B-4122"):
            p = "B-4122" + p
        if p not in seen:
            seen.add(p)
            cleaned.append(p)
    output_lines.append("; ".join(cleaned))

# final output
final = "\n".join(output_lines)
print(final)
