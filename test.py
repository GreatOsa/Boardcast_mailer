import pandas as pd

df = pd.read_excel("AuroraCreatorRegistration.xlsx")
df.columns = df.columns.astype(str).str.strip()

for name, post in zip(df["Full name"], df["Post"]):
    print(name)
    print(ascii(str(post)[:200]))
    print()