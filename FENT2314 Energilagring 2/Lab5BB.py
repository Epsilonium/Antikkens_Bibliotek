import xlrd
import numpy as np
import matplotlib.pyplot as plt

wb = xlrd.open_workbook('Lab5Black.xlsx')
sheet = wb.sheet_by_index(0)

area= 25

def newcurrent():
    current=[]
    current01 = 0.1/area
    current.append(current01)
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
    current75 = 7.5 / area
    current.append(current75)
    current10 = 10 / area
    current.append(current10)
    current155 = 15 / area
    current.append(current155)
    current20 = 20 / area
    current.append(current20)
    current255 = 25 / area
    current.append(current255)
    current30 = 30 / area
    current.append(current30)
    current35 = 35 / area
    current.append(current35)
    current40 = 40 / area
    current.append(current40)
    return current
  
def extractvoltage():
    voltage= []
    averagevolt =[]
    for i in range(sheet.nrows):
        if (i>=485) and (i<= 785): #0.1
            a = sheet.cell_value(i, 1)
            voltage.append(a)
        if (i==800):
            average = sum(voltage)/len(voltage)
            averagevolt.append(average)
            voltage = []
        elif (i>=1121) and (i<= 1421): #0.5
            a = sheet.cell_value(i, 1)
            voltage.append(a)
        if (i==1500):
            average = sum(voltage) / len(voltage)
            averagevolt.append(average)
            voltage = []
        elif (i>=1750) and (i<= 2050): # 1
            a = sheet.cell_value(i, 1)
            voltage.append(a)
        if (i==2200):
            average = sum(voltage) / len(voltage)
            averagevolt.append(average)
            voltage = []
        elif (i>=2376) and (i<= 2676): #1.5
            a = sheet.cell_value(i, 1)
            voltage.append(a)
        if (i==2800):
            average = sum(voltage) / len(voltage)
            averagevolt.append(average)
            voltage = []
        elif (i>=3007) and (i<= 3307): # 2
            a = sheet.cell_value(i, 1)
            voltage.append(a)
        if (i==3500):
            average = sum(voltage) / len(voltage)
            averagevolt.append(average)
            voltage = []
        elif (i >= 3636) and (i <= 3936):  # 2.5
            a = sheet.cell_value(i, 1)
            voltage.append(a)
        if (i==4500):
            average = sum(voltage) / len(voltage)
            averagevolt.append(average)
            voltage = []
        elif (i>=4875) and (i<= 5175): #5
            a = sheet.cell_value(i, 1)
            voltage.append(a)
        if (i==5500):
            average = sum(voltage) / len(voltage)
            averagevolt.append(average)
            voltage = []
        elif (i>=6119) and (i<= 6419): #7.5
            a = sheet.cell_value(i, 1)
            voltage.append(a)
        if (i==7000):
            average = sum(voltage) / len(voltage)
            averagevolt.append(average)
            voltage = []
        elif (i>=7712) and (i<= 8012): #10
            a = sheet.cell_value(i, 1)
            voltage.append(a)
        if (i==8500):
            average = sum(voltage) / len(voltage)
            averagevolt.append(average)
            voltage = []
        elif (i>=8950) and (i<= 9250): #15
            a = sheet.cell_value(i, 1)
            voltage.append(a)
        if (i==10000):
            average = sum(voltage) / len(voltage)
            averagevolt.append(average)
            voltage = []
        elif (i>=10229) and (i<= 10529): #20
            a = sheet.cell_value(i, 1)
            voltage.append(a)
        if (i==11000):
            average = sum(voltage) / len(voltage)
            averagevolt.append(average)
            voltage = []
        elif (i>=11531) and (i<= 11831): # 25
            a = sheet.cell_value(i, 1)
            voltage.append(a)
        if (i==12000):
            average = sum(voltage) / len(voltage)
            averagevolt.append(average)
            voltage = []
        elif (i>=12758) and (i<= 13058): #30
            a = sheet.cell_value(i, 1)
            voltage.append(a)
        if (i==13500):
            average = sum(voltage) / len(voltage)
            averagevolt.append(average)
            voltage = []
        elif (i>= 14073) and (i<= 14373): #35
            a = sheet.cell_value(i, 1)
            voltage.append(a)
        if (i==15000):
            average = sum(voltage) / len(voltage)
            averagevolt.append(average)
            voltage = []
        elif (i>=15331) and (i<= 15630): # 40
            a = sheet.cell_value(i, 1)
            voltage.append(a)
        if (i==15631):
            average = sum(voltage) / len(voltage)
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
