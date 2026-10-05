import pandas as pd

# Read Excel file
df = pd.read_excel("AuroraCreatorRegistration.xlsx")

# Get phone numbers
phone_numbers = df["WhatsApp number"].dropna().drop_duplicates()

# Save to TXT file
phone_numbers.to_csv("phone_numbers.txt", index=False, header=False)

print("✅ Phone numbers saved to phone_numbers.txt")