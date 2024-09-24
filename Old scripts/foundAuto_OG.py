import openpyxl
import pyautogui
import time

# Paths to workbooks and tracking file
jn_path = r"S:\AR Shared Job Files\Job Numbers.xlsx"

# Load workbook
try:
    jn_wb = openpyxl.load_workbook(jn_path)
except FileNotFoundError as e:
    print(f"Error: {e}")
    exit(1)

# Select sheet
try:
    jn_sheet = jn_wb["2024"]
except KeyError as e:
    print(f"Error: {e}")
    exit(1)

# Function to find the last row with data in JN
def find_last_row(sheet):
    for row in range(sheet.max_row, 0, -1):
        if sheet[f'A{row}'].value is not None:
            return row
    return None

# Get the last row with data in JN
last_row = find_last_row(jn_sheet)

if last_row is None:
    print("No rows with data found.")
    exit(0)

# Transfer data from JN to variables
Bx = jn_sheet[f'B{last_row}'].value
Cx = jn_sheet[f'C{last_row}'].value
Dx = jn_sheet[f'D{last_row}'].value
Ax = jn_sheet[f'A{last_row}'].value
Ex = jn_sheet[f'E{last_row}'].value
Fx = jn_sheet[f'F{last_row}'].value
Gx = jn_sheet[f'G{last_row}'].value
CAx = Fx * 0.9

# Logic for Division (CBx)
if Ax in ["EB", "ES", "AR", "PM"]:
    CBx = "Charlotte"
elif Ax in ["CO", "RV", "JD"]:
    CBx = "Raleigh"
else:
    CBx = ""

# Logic for Salesperson (CCx)
salesperson_map = {
    "CO": "ORT",
    "ES": "SMI2",
    "JD": "644",
    "AR": "426",
    "PM": "611",
    "EB": "EB",
    "RV": "512"
}

CCx = salesperson_map.get(Ax, "")

# Ensure the data is loaded correctly
print(f"B: {Bx}, C: {Cx}, D: {Dx}, A: {Ax}, E: {Ex}, F: {Fx}, G: {Gx}, CA: {CAx}, CB: {CBx}, CC: {CCx}")

# Coordinates mapping
coordinates = {
    'Bx': (2871, 179),
    'Bx_Cx': (2864, 218),
    'Dx': (2926, 246),
    'Ax': (2915, 271),
    'Ex_geo_area': (3014, 303),
    'project_class': (2958, 344),
    'Ex_payroll': (2955, 375),
    'Fx_contract': (3409, 300),
    'Fx_estimated_cost': (3424, 343),
    'change_tab1': (2756, 153),
    'Gx': (3494, 319),
    'change_tab2': (3406, 143),
    'CBx': (3064, 339),
    'change_tab3': (3731, 155),
    'CCx': (2886, 236)
}

# Allow time to open the correct window
print("You have 5 seconds to open the Foundation application window...")
time.sleep(5)

# Function to enter data into specific text boxes
def enter_data(data, position):
    pyautogui.click(position)
    pyautogui.typewrite(str(data))
    pyautogui.press('tab')

# Enter data into the Foundation application
enter_data(Bx, coordinates['Bx'])
time.sleep(1)  # Adjust sleep time as needed
enter_data(f"{Bx} {Cx}", coordinates['Bx_Cx'])
time.sleep(1)
enter_data(Dx, coordinates['Dx'])
time.sleep(1)
enter_data(Ax, coordinates['Ax'])
time.sleep(1)
enter_data(Ex, coordinates['Ex_geo_area'])
time.sleep(1)

# Project Class
pyautogui.click(coordinates['project_class'])
pyautogui.typewrite("COM")
pyautogui.press('tab')
time.sleep(1)

# Payroll Tax Group
enter_data(Ex, coordinates['Ex_payroll'])
time.sleep(1)

# Original Contract
enter_data(Fx, coordinates['Fx_contract'])
time.sleep(1)

# Original Estimated Cost
try:
    print(f"Calculated Estimated Cost: {CAx:.2f}")  # Debugging print
    enter_data(f"{CAx:.2f}", coordinates['Fx_estimated_cost'])
except Exception as e:
    print(f"Error calculating or entering estimated cost: {e}")

# Change Tab 1
pyautogui.click(coordinates['change_tab1'])
time.sleep(1)

# Order Date
enter_data(Gx, coordinates['Gx'])
time.sleep(1)

# Change Tab 2
pyautogui.click(coordinates['change_tab2'])
time.sleep(1)

# Division
enter_data(CBx, coordinates['CBx'])
time.sleep(1)

# Change Tab 3
pyautogui.click(coordinates['change_tab3'])
time.sleep(1)

# Salesperson
try:
    print(f"Entering Salesperson: {CCx} at position {coordinates['CCx']}")  # Debugging print
    enter_data(CCx, coordinates['CCx'])
except Exception as e:
    print(f"Error entering salesperson: {e}")
time.sleep(1)

print("Data entry completed.")

# Close workbook
jn_wb.close()
