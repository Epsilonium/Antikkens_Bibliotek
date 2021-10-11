import xlrd
import numpy as np
import matplotlib.pyplot as plt

wb = xlrd.open_workbook('lab1.xlsx')
sheet = wb.sheet_by_index(0)

def task7time():
    time = []
    for i in range(sheet.nrows):
        a = sheet.cell_value(i, 2)
        if (i >=10) and (i <= 497):
            time.append(a)
    return time

def task7current():
    current = []
    for i in range(sheet.nrows):
        a = sheet.cell_value(i, 6)
        if (i >=10) and (i <= 497):
            current.append(a)
    return current

def task7voltage():
    voltage = []
    for i in range(sheet.nrows):
        a = sheet.cell_value(i, 7)
        if (i >=10) and (i <= 497):
            voltage.append(a)
    return voltage

def capacity(a, b):
    products =[]
    for i,x in zip(a,b):
        products.append(i*x)

    return products

def task2time():
    time = []
    for i in range(sheet.nrows):
        a = sheet.cell_value(i, 2)
        if i >=498:
            time.append(a)
    return time

def task2current():
    current = []
    for i in range(sheet.nrows):
        a = sheet.cell_value(i, 6)
        if i >=498:
            current.append(a)
    return current

def task2voltage():
    voltage = []
    for i in range(sheet.nrows):
        a = sheet.cell_value(i, 7)
        if i >=498:
            voltage.append(a)
    return voltage

def task3current():
    current = []
    for i in range(sheet.nrows):
        a = sheet.cell_value(i, 9)
        if (i >=498) and (i <= 610):
            a=a*1000
            current.append(a)
    return current

def task3voltage():
    voltage = []
    for i in range(sheet.nrows):
        a = sheet.cell_value(i, 7)
        if (i >=498) and (i <= 610):
            voltage.append(a)
    return voltage

def task4current():
    current = []
    for i in range(sheet.nrows):
        a = sheet.cell_value(i, 10)
        if (i >=625) and (i <= 1685):
            a=a*1000
            current.append(a)
    return current

def task4voltage():
    voltage = []
    for i in range(sheet.nrows):
        a = sheet.cell_value(i, 7)
        if (i >=625) and (i <= 1685):
            voltage.append(a)
    return voltage

time = task2time()
current = task4current()
voltage = task4voltage()

"""
plt.plot(current, voltage, label='Voltage')
plt.xlabel('Current Capacity (mAh)')
plt.ylabel('Voltage (V)')
plt.legend()
plt.show()

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
#plt.plot(time, current, label ="Color", color= "green")
#plt.plot(time, voltage, label = "Voltage", color = "blue")
#plt.show()
