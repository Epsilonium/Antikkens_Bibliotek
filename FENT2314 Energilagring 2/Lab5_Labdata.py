import xlrd
import numpy as np
import matplotlib.pyplot as plt

wb = xlrd.open_workbook('Labresultater.xlsx')
sheet = wb.sheet_by_index(0)

area= 25

def newcurrent():
    current=[]
    current0 = 0/area
    current.append(current0)
    current05 = 0.5 / area
    current.append(current05)
    current1 = 1 / area
    current.append(current1)
    current15 = 1.5 / area
    current.append(current15)
    current2 = 2 / area
    current.append(current2)
    current25 = 2.5 / area
    current.append(current25)
    current5 = 5 / area
    current.append(current5)
    return current

def extractvoltage():
    voltage= []
    averagevolt =[]
    for i in range(sheet.nrows):
        if (i>=218) and (i<= 518): #0.1
            a = sheet.cell_value(i, 1)
            voltage.append(a)
        if (i==800):
            average = sum(voltage)/len(voltage)
            averagevolt.append(average)
            voltage = []
        elif (i>=1112) and (i<= 1412): #0.5
            a = sheet.cell_value(i, 1)
            voltage.append(a)
        if (i==1500):
            average = sum(voltage) / len(voltage)
            averagevolt.append(average)
            voltage = []
        elif (i>=1821) and (i<= 2121): # 1
            a = sheet.cell_value(i, 1)
            voltage.append(a)
        if (i==2200):
            average = sum(voltage) / len(voltage)
            averagevolt.append(average)
            voltage = []
        elif (i>=2485) and (i<= 2785): #1.5
            a = sheet.cell_value(i, 1)
            voltage.append(a)
        if (i==2800):
            average = sum(voltage) / len(voltage)
            averagevolt.append(average)
            voltage = []
        elif (i>=3171) and (i<= 3471): # 2
            a = sheet.cell_value(i, 1)
            voltage.append(a)
        if (i==3500):
            average = sum(voltage) / len(voltage)
            averagevolt.append(average)
            voltage = []
        elif (i >= 3840) and (i <= 4140):  # 2.5
            a = sheet.cell_value(i, 1)
            voltage.append(a)
        if (i==4500):
            average = sum(voltage) / len(voltage)
            averagevolt.append(average)
            voltage = []
        elif (i>=5557) and (i<= 5857): #5
            a = sheet.cell_value(i, 1)
            voltage.append(a)
        if (i==5858):
            average = sum(voltage) / len(voltage)
            print(average)
            averagevolt.append(average)
            voltage = []
    return averagevolt

def effect(a,b):
    products = []
    for num1, num2 in zip(a,b):
        products.append(num1*num2)
    return products
current = newcurrent()
voltage = extractvoltage()
power = effect(current, voltage)

"""
plt.scatter(current, voltage, label='Voltage')
plt.xlabel('Current Density (A/cm^2)')
plt.ylabel('Voltage (V)')
plt.legend()
plt.show()
"""
plt.scatter(current, power, label='Power')
plt.xlabel('Current Density (A/cm^2)')
plt.ylabel('Power (W/cm^2)')
plt.legend()
plt.show()
