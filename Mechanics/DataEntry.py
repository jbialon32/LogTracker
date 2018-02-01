'''
Created on Jan 26, 2018

@author: James
'''

from openpyxl import load_workbook

def NumEntry():
    file = open("Mechanics/MileageTracker.txt", "r")
    storedMileage = file.readline()
    file.close()
    
    storedMileage = int(storedMileage)
    
    goodDays = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
    goodWeeks = ["1", "2", "3", "4"]
    empList = []
    empNum = 1
    
    numbersBook = load_workbook("Mechanics/Log.xlsx")
    
    week = str(input("Enter week of period(1-4): "))
    day = str(input("\nEnter day of the week: "))
    
    day = day.lower()

    if week in goodWeeks:        
        if day in goodDays:                    
            try:
                mileage = int(input("\nEnter ending milage: "))
                numStops = int(input("\nEnter total stops: "))
            except:
                print("ERROR: Mileage and total stops must be integers (whole numbers).")
            
            try:
                gasCost = float(input("\nEnter gas costs: "))
            except:
                print("ERROR: Gas costs must be a floating point value (Ex. 3.56")    
                
            try:
                numEmployees = int(input("\nHow many employees were on the clock: "))
            except:
                print("ERROR: Number of employees must be integers (whole numbers).")
            
            while empNum <= numEmployees:
                empName = str(input("\nEnter employee #" + str(empNum) + "'s name: "))
                empList.append(empName)
                empNum += 1
                
            nameString = ", ".join(empList)
            nameString = nameString.title()
            
            if week == "1":
                weekSheet = numbersBook["week1"]
                    
            if week == "2":
                weekSheet = numbersBook["week2"]
                
            if week == "3":
                weekSheet = numbersBook["week3"]
            
            if week == "4":
                weekSheet = numbersBook["week4"]
                
            if day == "thursday":
                weekSheet["B2"] = storedMileage
                weekSheet["B3"] = mileage
                weekSheet["B5"] = gasCost
                weekSheet["B6"] = numStops
                weekSheet["B7"] = nameString
                numbersBook.save("Mechanics/Log.xlsx")
                storedMileage = mileage
                file = open("Mechanics/MileageTracker.txt", "w")
                file.write(str(storedMileage))
                file.close()
                    
            if day == "friday":
                weekSheet["C2"] = storedMileage
                weekSheet["C3"] = mileage
                weekSheet["C5"] = gasCost
                weekSheet["C6"] = numStops
                weekSheet["C7"] = nameString
                numbersBook.save("Mechanics/Log.xlsx")
                storedMileage = mileage
                file = open("Mechanics/MileageTracker.txt", "w")
                file.write(str(storedMileage))
                file.close()
                    
            if day == "saturday":
                weekSheet["D2"] = storedMileage
                weekSheet["D3"] = mileage
                weekSheet["D5"] = gasCost
                weekSheet["D6"] = numStops
                weekSheet["D7"] = nameString
                numbersBook.save("Mechanics/Log.xlsx")
                storedMileage = mileage
                file = open("Mechanics/MileageTracker.txt", "w")
                file.write(str(storedMileage))
                file.close()
                    
            if day == "sunday":
                weekSheet["E2"] = storedMileage
                weekSheet["E3"] = mileage
                weekSheet["E5"] = gasCost
                weekSheet["E6"] = numStops
                weekSheet["E7"] = nameString
                numbersBook.save("Mechanics/Log.xlsx")
                storedMileage = mileage
                file = open("Mechanics/MileageTracker.txt", "w")
                file.write(str(storedMileage))
                file.close()
                    
            if day == "monday":
                weekSheet["F2"] = storedMileage
                weekSheet["F3"] = mileage
                weekSheet["F5"] = gasCost
                weekSheet["F6"] = numStops
                weekSheet["F7"] = nameString
                numbersBook.save("Mechanics/Log.xlsx")
                storedMileage = mileage
                file = open("Mechanics/MileageTracker.txt", "w")
                file.write(str(storedMileage))
                file.close()
                    
            if day == "tuesday":
                weekSheet["G2"] = storedMileage
                weekSheet["G3"] = mileage
                weekSheet["G5"] = gasCost
                weekSheet["G6"] = numStops
                weekSheet["G7"] = nameString
                numbersBook.save("Mechanics/Log.xlsx")
                storedMileage = mileage
                file = open("Mechanics/MileageTracker.txt", "w")
                file.write(str(mileage))
                file.close()
                    
            if day == "wednesday":
                weekSheet["H2"] = storedMileage
                weekSheet["H3"] = mileage
                weekSheet["H5"] = gasCost
                weekSheet["H6"] = numStops
                weekSheet["H7"] = nameString
                numbersBook.save("Mechanics/Log.xlsx")
                storedMileage = mileage
                file = open("Mechanics/MileageTracker.txt", "w")
                file.write(str(storedMileage))
                file.close()
                    
        else:
            print("enter valid day. (Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Sunday)")    
        
    else:
        print("Period week must be in range (1, 2, 3, 4")
        
    