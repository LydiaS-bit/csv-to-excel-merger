import pandas as pd
desktop_path = r'C:\Users\lee\OneDrive\Desktop\\'  
try:
    customers = pd.read_csv(desktop_path + 'customers.csv')
    print("Customers CSV loaded successfully.")
except Exception as e:
    print(f"Error loading customers CSV: {e}")
try:
    delivery_issues = pd.read_csv(desktop_path + 'delivery_issues.csv')
    print("Delivery Issues CSV loaded successfully.")
except Exception as e:
    print(f"Error loading delivery issues CSV: {e}")
try:
    orders = pd.read_csv(desktop_path + 'orders.csv')
    print("Orders CSV loaded successfully.")
except Exception as e:
    print(f"Error loading orders CSV: {e}")
try:
    weather = pd.read_csv(desktop_path + 'weather.csv')
    print("Weather CSV loaded successfully.")
except Exception as e:
    print(f"Error loading weather CSV: {e}")
try:
    with pd.ExcelWriter(desktop_path + 'merged_csvs.xlsx', engine='openpyxl') as writer:
        customers.to_excel(writer, sheet_name='Customers', index=False)
        delivery_issues.to_excel(writer, sheet_name='Delivery Issues', index=False)
        orders.to_excel(writer, sheet_name='Orders', index=False)
        weather.to_excel(writer, sheet_name='Weather', index=False)
    print("CSV files have been merged into 'merged_csvs.xlsx'!")
except Exception as e:
    print(f"Error writing to Excel: {e}")
