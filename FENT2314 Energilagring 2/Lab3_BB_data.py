import xlrd
import numpy as np
import matplotlib.pyplot as plt

wb = xlrd.open_workbook('Cell_data.xlsx')
sheet = wb.sheet_by_index(0)

two = abs(sheet.cell_value(355, 3))*1/3600

def task5time():
    time = []
    for i in range(sheet.nrows):
        a = sheet.cell_value(i, 2)
        if i != 0:
            time.append(a)
    return time

def task5current():
    current = []
    for i in range(sheet.nrows):
        a = sheet.cell_value(i, 3)
        if i != 0:
            current.append(a)
    return current

def task5voltage():
    voltage = []
    for i in range(sheet.nrows):
        a = sheet.cell_value(i, 4)
        if i != 0:
            voltage.append(a)
    return voltage

def task6voltage():
    voltage = []
    for i in range(sheet.nrows):
        a = sheet.cell_value(i, 4)
        if i >= 354:
            voltage.append(a)
    return voltage

def task6current(a, b):
    old_d = abs(sheet.cell_value(355, 3))*1/3600
    products =[]
    for i in range(sheet.nrows):
        if (i >= 354):
            a =sheet.cell_value(i, 2)
            b = sheet.cell_value((i-1), 2)
            delta_t= (a-b)/3600
            c = abs(sheet.cell_value(i, 4))
            current_d = ((delta_t*c) + old_d)
            products.append(current_d)
            old_d = current_d
    return products

time = task5time()
current = task5current()
voltage = task5voltage()
dischargevolt = task6voltage()

AmpereHour = task6current(time, current)

"""
fig, ax1 = plt.subplots()

ax1.set_xlabel('time (s)')
ax1.set_ylabel('Voltage', color="blue")
ax1.plot(time, voltage, color = "blue", label = "Voltage")
ax1.tick_params(axis='y', labelcolor="blue")

ax2 = ax1.twinx()

ax2.set_ylabel('Current', color= "red")
ax2.plot(time, current, color = "red", label="Current")
ax2.tick_params(axis = 'y', labelcolor = "red")

fig.tight_layout()
plt.show()
"""
plt.plot(AmpereHour, dischargevolt, label='Voltage')
plt.xlabel('Current Capacity (mAh)')
plt.ylabel('Voltage (V)')
plt.legend()
plt.show()

