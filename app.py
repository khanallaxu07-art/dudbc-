<<<<<<< HEAD
import pandas as pd

# Load DUDBC/DOR Excel
df = pd.read_excel("dudbc_dor.xlsx")  # or CSV: pd.read_csv("dudbc_dor.csv")

def calculate_material(item_name, quantity):
    """Calculate material for a single item"""
    item_row = df[df['Item'].str.lower() == item_name.lower()]
    if item_row.empty:
        return None
    item_data = item_row.iloc[0]
    breakdown = {}
    for col in df.columns:
        if col != 'Item' and pd.notna(item_data[col]):
            breakdown[col] = item_data[col] * quantity
    return breakdown

def main():
    print("Available items:", ", ".join(df['Item'].tolist()))
    order_list = []

    # Accept multiple items
    while True:
        item_name = input("\nEnter item name (or 'done' to finish): ").strip()
        if item_name.lower() == 'done':
            break
        try:
            quantity = float(input(f"Enter quantity for '{item_name}' (m3/m2/etc.): "))
        except ValueError:
            print("Invalid quantity, try again.")
            continue
        order_list.append((item_name, quantity))

    if not order_list:
        print("No items entered. Exiting.")
        return

    # Calculate breakdown per item
    breakdown_list = []
    for item_name, quantity in order_list:
        breakdown = calculate_material(item_name, quantity)
        if breakdown is None:
            print(f"Warning: Item '{item_name}' not found. Skipping.")
            continue
        breakdown['Item'] = item_name
        breakdown['Quantity'] = quantity
        breakdown_list.append(breakdown)

    if not breakdown_list:
        print("No valid items to calculate. Exiting.")
        return

    breakdown_df = pd.DataFrame(breakdown_list)

    # Calculate total material requirement
    total_materials = breakdown_df.drop(columns=['Item', 'Quantity']).sum().to_frame(name='Total')
    total_materials = total_materials.reset_index().rename(columns={'index': 'Material'})

    # Display results
    print("\n=== Material Breakdown per Item ===")
    print(breakdown_df.to_string(index=False))
    print("\n=== Total Material Requirement ===")
    print(total_materials.to_string(index=False))

    # Save to Excel
    output_file = "material_breakdown_result.xlsx"
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        breakdown_df.to_excel(writer, sheet_name='Item Breakdown', index=False)
        total_materials.to_excel(writer, sheet_name='Total Materials', index=False)
    print(f"\nResults saved to '{output_file}'")

if __name__ == "__main__":
    main()
=======
import pandas as pd

# Load DUDBC/DOR Excel
df = pd.read_excel("dudbc_dor.xlsx")  # or CSV: pd.read_csv("dudbc_dor.csv")

def calculate_material(item_name, quantity):
    """Calculate material for a single item"""
    item_row = df[df['Item'].str.lower() == item_name.lower()]
    if item_row.empty:
        return None
    item_data = item_row.iloc[0]
    breakdown = {}
    for col in df.columns:
        if col != 'Item' and pd.notna(item_data[col]):
            breakdown[col] = item_data[col] * quantity
    return breakdown

def main():
    print("Available items:", ", ".join(df['Item'].tolist()))
    order_list = []

    # Accept multiple items
    while True:
        item_name = input("\nEnter item name (or 'done' to finish): ").strip()
        if item_name.lower() == 'done':
            break
        try:
            quantity = float(input(f"Enter quantity for '{item_name}' (m3/m2/etc.): "))
        except ValueError:
            print("Invalid quantity, try again.")
            continue
        order_list.append((item_name, quantity))

    if not order_list:
        print("No items entered. Exiting.")
        return

    # Calculate breakdown per item
    breakdown_list = []
    for item_name, quantity in order_list:
        breakdown = calculate_material(item_name, quantity)
        if breakdown is None:
            print(f"Warning: Item '{item_name}' not found. Skipping.")
            continue
        breakdown['Item'] = item_name
        breakdown['Quantity'] = quantity
        breakdown_list.append(breakdown)

    if not breakdown_list:
        print("No valid items to calculate. Exiting.")
        return

    breakdown_df = pd.DataFrame(breakdown_list)

    # Calculate total material requirement
    total_materials = breakdown_df.drop(columns=['Item', 'Quantity']).sum().to_frame(name='Total')
    total_materials = total_materials.reset_index().rename(columns={'index': 'Material'})

    # Display results
    print("\n=== Material Breakdown per Item ===")
    print(breakdown_df.to_string(index=False))
    print("\n=== Total Material Requirement ===")
    print(total_materials.to_string(index=False))

    # Save to Excel
    output_file = "material_breakdown_result.xlsx"
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        breakdown_df.to_excel(writer, sheet_name='Item Breakdown', index=False)
        total_materials.to_excel(writer, sheet_name='Total Materials', index=False)
    print(f"\nResults saved to '{output_file}'")

if __name__ == "__main__":
    main()
>>>>>>> 4d4179d (Fix requirements.txt for Render deployment)
