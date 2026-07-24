import json
from pathlib import Path

# Load your vision cache files
cache_dir = Path(r"D:\0-Batch-AWARDS\processed_files\Review\Vision_Cache") # adjust path
cache_files = list(cache_dir.glob("*.json"))

print(f"Total Vision Cache Entries: {len(cache_files)}")

# Audit sample entries for detail depth
short_responses = 0
for cf in cache_files:
    with open(cf, "r") as f:
        data = json.load(f)
        # Check output length or token count as a proxy for detail
        output_text = data.get("response", "") or data.get("text", "")
        if len(output_text.split()) < 30: # unusually brief
            short_responses += 1

print(f"Potentially low-detail vision extractions (<30 words): {short_responses}")