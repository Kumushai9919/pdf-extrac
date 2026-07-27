import pprint, os
import fitz
from main import extract_records
os.makedirs("debug_images", exist_ok=True)
doc = fitz.open("S1.pdf")
recs, skipped = extract_records(doc, "debug_images", dpi=300, img_dpi=300, template_name="S1")
pprint.pprint(recs[:6])
print("skipped:", skipped)
