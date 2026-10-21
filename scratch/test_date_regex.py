import re

DATE_RANGE_REGEX = re.compile(
    r'(?:'
    r'(?:(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+)?(?:19|20)\d{2}\s*[-–—/]\s*'
    r'(?:(?:(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+)?(?:19|20)\d{2}|PRESENT|CURRENT|ONGOING|TODAY)'
    r'|'
    r'(?:19|20)\d{2}\s*[-–—/]\s*(?:19|20)\d{2}'
    r')',
    re.IGNORECASE
)

test_cases = [
    "Borcelle Studio | 2030 - PRESENT\nMarketing Manager & Specialist",
    "Fauget Studio | 2025 - 2029\nMarketing Manager & Specialist",
    "Studio Shodwe | 2024 - 2025\nMarketing Manager & Specialist",
    "Borcelle Studio | 2030 - PRESENT\nMarketing Manager & Specialist",
    "Fauget Studio | 2025 - 2029\nMarketing Specialist",
    "Senior Designer at Google (2021 - 2023)\nLead UI/UX",
    "WORK EXPERIENCE",
    "PROFILE",
    "Lorem ipsum dolor sit amet"
]

for tc in test_cases:
    m = DATE_RANGE_REGEX.search(tc)
    if m:
        date_str = m.group(0).strip()
        clean = tc[:m.start()] + tc[m.end():]
        clean = re.sub(r'[\(\)]', '', clean)
        clean = re.sub(r'\s*\|\s*', '\n', clean)
        lines = [l.strip(' \t,-|') for l in clean.split('\n') if l.strip(' \t,-|')]
        print(f"MATCH: {tc!r}")
        print(f"  Lines: {lines}")
        print(f"  Date:  {date_str}")
    else:
        print(f"NO MATCH: {tc!r}")

