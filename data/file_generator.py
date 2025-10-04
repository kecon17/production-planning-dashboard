#%%
# Script to generate sample Excel files for production planning
# Files: demand.xlsx, stations.xlsx, products.xlsx, times.xlsx, training.xlsx
import pandas as pd
import random
import string

# Configuration
NUM_PRODUCTS = 6
NUM_OPERATORS = 25
MIN_SUBASSEMBLIES = 10
MAX_SUBASSEMBLIES = 30
MIN_MONTHLY_DEMAND = 20
MAX_MONTHLY_DEMAND = 50
YEAR = 2025

# Sample data for generating realistic values
product_prefixes = ['Robs-Std', 'Robs-Pro', 'MT-Beta', 'MT-Delta', 'FlowTX-M', 'FlowTX-L']
module_types = ['Bottom Structure', 'Top Structure', 'Electronics', 'Covers', 'Vacuum Pump', 'Pneumatic System', 'Mixer',
                'Front Panel', 'Thermal Chamber', 'Gripper', 'Dispenser', 'Optical System', 'PC']
technologies = ['Mechanical', 'Robotics', 'Fluidics', 'Packaging', 'Optics', 'Electronics', 'Pneumatics', 'Hydraulics']
operator_first_names = ['Arnauld', 'Berni', 'Clara', 'David', 'Emma', 'Frank', 'Gloria', 'Hans',
                        'Iris', 'Joan', 'Karl', 'Laura', 'Markus', 'Nina', 'Oscar', 'Paula',
                        'Quim', 'Rosa', 'Sergi', 'Tesa', 'Ulrich', 'Victor', 'Wendy', 'Xavier', 'Yolanda']
operator_last_names = ['Girald', 'Armend', 'Fontana', 'Lopez', 'Lloyd', 'Martinez', 'Garcia',
                       'Schamg', 'Soler', 'Font', 'Vila', 'Mas', 'Costa', 'Da Costa', 'Sanz']

# Generate product data
products = []
for i in range(NUM_PRODUCTS):
    code = ''.join(random.choices(string.digits + string.ascii_lowercase, k=random.randint(3, 5)))
    products.append({
        'code': code,
        'description': product_prefixes[i]
    })

# Generate operator data
operators = []
for i in range(NUM_OPERATORS):
    first = random.choice(operator_first_names)
    last = random.choice(operator_last_names)
    user_code = first[:1].lower() + last.replace(' ', '')[:2].lower() + str(random.randint(1, 9))
    operators.append({
        'code': user_code,
        'name': f"{first} {last}",
        'short_name': first[:4]  # NomCurt = first 4 letters of first name
    })

# Generate stations (UT)
stations = [f"UT-{str(i).zfill(5)}" for i in range(1, 16)]

# ============================================
# 1. Generate demand.xlsx
# ============================================
demand_data = []
for product in products:
    for month in range(1, 13):
        base_demand = random.randint(MIN_MONTHLY_DEMAND, MAX_MONTHLY_DEMAND)
        seasonal_factor = 1 + 0.2 * ((month - 6.5) / 6.5)
        units = int(base_demand * seasonal_factor)

        demand_data.append({
            'Any': YEAR,
            'Mes': month,
            'CodiProjecte': product['code'],
            'DescripcioProjecte': product['description'],
            'Unitats': units
        })

df_demand = pd.DataFrame(demand_data)

# ============================================
# 2. Generate stations.xlsx
# ============================================
stations_data = []
all_modules = []

for product in products:
    num_subassemblies = random.randint(MIN_SUBASSEMBLIES, MAX_SUBASSEMBLIES)

    for i in range(num_subassemblies):
        module_code = str(random.randint(600000, 999999))
        prefix = random.choice(['Bottom', 'Top', 'Electronics', 'Covers', 'Vacuum Pump', 'Pneumatic System', 'Mixer',
                                'Front Panel', 'Thermal Chamber', 'Gripper', 'Dispenser', 'Optical System', 'PC'])
        suffix = random.choice(module_types)
        module_desc = f"{prefix}{suffix}"
        station = random.choice(stations)
        technology = random.choice(technologies)

        module_info = {
            'CodiProjecte': product['code'],
            'DescripcioProjecte': product['description'],
            'CodiModul': module_code,
            'DescripcioModul': module_desc,
            'UT': station,
            'TecnologiaPerUT': technology
        }

        stations_data.append(module_info)
        all_modules.append(module_info)

df_stations = pd.DataFrame(stations_data)

# ============================================
# 3. Generate times.xlsx
# ============================================
times_data = []
for module in all_modules:
    standard_time = round(random.uniform(2, 12), 2)
    times_data.append({
        'CodiProjecte': module['CodiProjecte'],
        'DescripcioProjecte': module['DescripcioProjecte'],
        'CodiModul': module['CodiModul'],
        'DescripcioModul': module['DescripcioModul'],
        'TempsEstandar': standard_time
    })

df_times = pd.DataFrame(times_data)

# ============================================
# 4. Generate training.xlsx (with NomCurt)
# ============================================
training_data = []
for module in all_modules:
    num_trained = random.randint(int(NUM_OPERATORS * 0.3), int(NUM_OPERATORS * 0.7))
    trained_operators = random.sample(operators, num_trained)

    for operator in trained_operators:
        training_data.append({
            'CodiProjecte': module['CodiProjecte'],
            'DescripcioProjecte': module['DescripcioProjecte'],
            'CodiModul': module['CodiModul'],
            'DescripcioModul': module['DescripcioModul'],
            'Usuari': operator['code'],
            'Nom': operator['name'],
            'NomCurt': operator['short_name']
        })

df_training = pd.DataFrame(training_data)

# ============================================
# 5. Generate products.xlsx
# ============================================
products_data = []
for module in all_modules:
    # ModulCode = abbreviation = first letters of each word in description
    words = module['DescripcioModul'].replace('-', ' ').split()
    modul_code = ''.join([w[0].upper() for w in words])

    # ModulPerProduct = random quantity of this module per product (2–10)
    modul_qty = random.randint(2, 10)

    products_data.append({
        'CodiProjecte': module['CodiProjecte'],
        'DescripcioProjecte': module['DescripcioProjecte'],
        'CodiModul': module['CodiModul'],
        'DescripcioModul': module['DescripcioModul'],
        'ModulCode': modul_code,
        'ModulPerProduct': modul_qty
    })

df_products = pd.DataFrame(products_data)

# ============================================
# Save all files
# ============================================
with pd.ExcelWriter('demand.xlsx', engine='openpyxl') as writer:
    df_demand.to_excel(writer, sheet_name='demand', index=False)

with pd.ExcelWriter('stations.xlsx', engine='openpyxl') as writer:
    df_stations.to_excel(writer, sheet_name='stations', index=False)

with pd.ExcelWriter('times.xlsx', engine='openpyxl') as writer:
    df_times.to_excel(writer, sheet_name='times', index=False)

with pd.ExcelWriter('training.xlsx', engine='openpyxl') as writer:
    df_training.to_excel(writer, sheet_name='training', index=False)

with pd.ExcelWriter('products.xlsx', engine='openpyxl') as writer:
    df_products.to_excel(writer, sheet_name='products', index=False)

print("Files generated successfully!")
print(f"\n=== Summary ===")
print(f"Products: {len(products)}")
print(f"Operators: {len(operators)}")
print(f"Total subassemblies: {len(all_modules)}")
print(f"Demand records: {len(df_demand)}")
print(f"Stations records: {len(df_stations)}")
print(f"Times records: {len(df_times)}")
print(f"Training records: {len(df_training)}")
print(f"\nProduct codes: {[p['code'] for p in products]}")
print(f"\nFiles created:")
print("  - demand.xlsx")
print("  - stations.xlsx")
print("  - times.xlsx")
print("  - training.xlsx")
print("  - products.xlsx")

#%%
print(df_demand.head())
print(df_stations.head())
print(df_times.head())
print(df_training.head())
print(df_products.head())

