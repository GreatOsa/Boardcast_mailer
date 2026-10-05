import pandas as pd

df = pd.read_excel("AuroraCreatorRegistration.xlsx")

# Clean email column
df["Email"] = df["Email"].astype(str).str.strip().str.lower()

# Count emails
email_counts = df["Email"].value_counts()

# Get emails appearing more than once
duplicates = email_counts[email_counts > 1]

if duplicates.empty:
    print("✅ No duplicate emails found.")
else:
    print("⚠️ Duplicate emails found:\n")
    print(duplicates)