"""Create a small, fake, messy sample file to demo the cleaner.

All names and values are made up. Run:  python make_sample.py
"""
import pandas as pd

rows = [
    ["  alice johnson ", "Alice.Johnson@Example.com", "10/05/2026", "Toronto", "1,200.50", "Paid"],
    ["BOB SMITH", "bob.smith@example.com", "Oct 7 2026", "ottawa", "$850", "paid"],
    ["carol Lee", "carol.lee@example.com", "2026-10-09", "Montreal ", "420", "PENDING"],
    ["  alice johnson ", "Alice.Johnson@Example.com", "10/05/2026", "Toronto", "1,200.50", "Paid"],
    ["David Chen", None, "11/10/2026", "Vancouver", "975.25", "Pending"],
    ["emma  Wilson", "emma.wilson@example.com", "Oct 12 2026", None, "$1,500", "Paid"],
    ["BOB SMITH", "bob.smith@example.com", "Oct 7 2026", "ottawa", "$850", "paid"],
    ["Frank Moore", "frank.moore@example.com", "2026/10/15", "Calgary", None, "Overdue"],
    ["grace HOPPER", "grace.h@example.com", "15-10-2026", "toronto", "300", "paid"],
    ["Henry Ford", "henry.ford@example.com", "Oct 18 2026", "Ottawa", "2,050", "Pending"],
    ["  ivy Zhang", "IVY.ZHANG@example.com", "2026-10-20", "Montreal", "615.75", "PAID"],
    ["Jack Brown", "jack.brown@example.com", "not available", "Halifax", "$99.99", "Overdue"],
]

df = pd.DataFrame(
    rows, columns=["Customer Name", "Email", "Order Date", "City", "Amount", "Status"]
)
df.to_excel("messy_sample.xlsx", index=False)
print(f"Wrote messy_sample.xlsx with {len(df)} rows")
