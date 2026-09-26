import pandas as pd
import random
import numpy as np

print("Generating data...")

# Sample base data
names = ["Rahul", "Amit", "Priya", "Neha", "Vikram", "Rohan", "Suresh", "Anjali", "Karan", "Pooja"]
states = ["Delhi", "Maharashtra", "Gujarat", "Punjab", "Haryana", "Karnataka", "Bengal", "Rajasthan", "UP", "MP"]

data = []

# Generate 1,000,00 rows (1 Lakh) - Aap isko 1000000 (10 Lakh) bhi kar sakte hain
for i in range(1000000):
    name = random.choice(names) + f" {random.randint(1, 1000)}"
    
    # 70% valid phone (10 digits), 20% invalid (8-9 digits), 10% blank
    phone_chance = random.random()
    if phone_chance < 0.7:
        phone = f"9{random.randint(100000000, 999999999)}" # 10 digits
    elif phone_chance < 0.9:
        phone = f"{random.randint(1000000, 99999999)}" # 8 or 9 digits (INVALID)
    else:
        phone = "" # BLANK
        
    state = random.choice(states)
    
    # Randomly make some Names or States blank
    if random.random() < 0.05: name = ""
    if random.random() < 0.05: state = ""

    data.append([name, phone, state])

df = pd.DataFrame(data, columns=["Name", "Phone", "State"])

# Add some exact duplicates intentionally
print("Adding duplicates...")
duplicates = df.sample(n=5000) # 5000 rows ko duplicate kar diya
df = pd.concat([df, duplicates], ignore_index=True)

# Save to Excel
print("Saving to Excel... (This might take a minute)")
df.to_excel("bulk_test_data.xlsx", index=False)
print("File 'bulk_test_data.xlsx' created successfully!")