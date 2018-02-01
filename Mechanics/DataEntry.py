'''
Created on Jan 26, 2018

@author: James
'''

from openpyxl import load_workbook

def NumEntry():
    
    logLocation = "Mechanics/Log.xlsx"
    mileageLocation = "Mechanics/MileageTracker.txt"
    
    # opens text file containing previous days mileage.
    file = open(mileageLocation, "r")
    storedMileage = file.readline()
    file.close()
    
    storedMileage = int(storedMileage)
    
    # Create "checks"
    goodDays = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
    goodWeeks = ["1", "2", "3", "4"]
    gottenMiles = False
    gottenStops = False
    gasGotten = False
    gottenNumEmp = False
    empList = []
    empNum = 1
    
    # Load in Excel workbook
    numbersBook = load_workbook(logLocation)
    
    # Create infinite loop so we can reset said loop as needed
    while True:        
        
        # Gets input for week used for workbook page info    
        week = str(input("Enter week of period(1-4): "))
        
        # Makes sure users input is recognizable by the program or resets loop
        if week in goodWeeks:
            
            # Checks week info in more detail to select correct workbook page
            if week == "1":
                    
                weekSheet = numbersBook["week1"]
                        
            if week == "2":
                    
                weekSheet = numbersBook["week2"]
                    
            if week == "3":
                    
                weekSheet = numbersBook["week3"]
                
            if week == "4":
                    
                weekSheet = numbersBook["week4"]
            
            # Gets input for week used for workbook column info 
            day = str(input("\nEnter day of the week: "))
            
            # Casts day as lower for easier translation
            day = day.lower()
            
            # Makes sure users input is recognizable by the program or resets loop        
            if day in goodDays: 
                
                if day == "thursday":
                
                    # Does a check if overwriting existing data
                    if weekSheet["B2"].value != None:
                        safetyFirst = False
                        print("\nWARNING: Data already exists for this day!")
                        while safetyFirst == False:
                            confirmation = input("Continue? (Y/N)")
                            confirmation = confirmation.capitalize()
                            if confirmation == "Y" or confirmation == "YES":
                                safetyFirst = True
                            else:
                                exit()
                                
                if day == "friday":
                
                    if weekSheet["C2"].value != None:
                        safetyFirst = False
                        print("\nWARNING: Data already exists for this day!")
                        while safetyFirst == False:
                            confirmation = input("Continue? (Y/N)")
                            confirmation = confirmation.capitalize()
                            if confirmation == "Y" or confirmation == "YES":
                                safetyFirst = True
                            else:
                                exit()
                                
                if day == "saturday":
                
                    if weekSheet["D2"].value != None:
                        safetyFirst = False
                        print("\nWARNING: Data already exists for this day!")
                        while safetyFirst == False:
                            confirmation = input("Continue? (Y/N)")
                            confirmation = confirmation.capitalize()
                            if confirmation == "Y" or confirmation == "YES":
                                safetyFirst = True
                            else:
                                exit()
                                
                if day == "sunday":

                    if weekSheet["E2"].value != None:
                        safetyFirst = False
                        print("\nWARNING: Data already exists for this day!")
                        while safetyFirst == False:
                            confirmation = input("Continue? (Y/N)")
                            confirmation = confirmation.capitalize()
                            if confirmation == "Y" or confirmation == "YES":
                                safetyFirst = True
                            else:
                                exit()
                                
                if day == "monday":

                    if weekSheet["F2"].value != None:
                        safetyFirst = False
                        print("\nWARNING: Data already exists for this day!")
                        while safetyFirst == False:
                            confirmation = input("Continue? (Y/N)")
                            confirmation = confirmation.capitalize()
                            if confirmation == "Y" or confirmation == "YES":
                                safetyFirst = True
                            else:
                                exit()
                                
                if day == "tuesday":

                    if weekSheet["G2"].value != None:
                        safetyFirst = False
                        print("\nWARNING: Data already exists for this day!")
                        while safetyFirst == False:
                            confirmation = input("Continue? (Y/N)")
                            confirmation = confirmation.capitalize()
                            if confirmation == "Y" or confirmation == "YES":
                                safetyFirst = True
                            else:
                                exit()
                                
                if day == "wednesday":

                    if weekSheet["H2"].value != None:
                        safetyFirst = False
                        print("\nWARNING: Data already exists for this day!")
                        while safetyFirst == False:
                            confirmation = input("Continue? (Y/N)")
                            confirmation = confirmation.capitalize()
                            if confirmation == "Y" or confirmation == "YES":
                                safetyFirst = True
                            else:
                                exit()
                
                # Creates a loop to make sure user successfully enters info
                while gottenMiles == False:  
                                      
                    try:
                        
                        mileage = int(input("\nEnter ending mileage: "))
                        gottenMiles = True
                        
                    except:
                        
                        # Makes sure user enters correct data
                        print("ERROR: Mileage must be an integer (whole number).")
                        continue
                
                while gottenStops == False:
                    
                    try:
                        
                        numStops = int(input("\nEnter total stops: "))
                        gottenStops = True
                        
                    except:
                        
                        print("ERROR: Total stops must be an integer (whole number).")
                        continue
                
                while gasGotten == False:
                
                    try:
                    
                        gasCost = float(input("\nEnter gas costs: "))
                    
                        while gasCost < 0:
                            print("Please enter a positive integer or 0")
                            gasCost = float(input("\nEnter gas costs: "))  
                        
                        gasGotten = True
                        
                    except:
                        
                        print("ERROR: Gas costs must be a floating point value (Ex. 3.56")
                        continue
                    
                while gottenNumEmp == False:
                    
                    try:
                        
                        numEmployees = int(input("\nHow many employees were on the clock: "))
                        
                        while numEmployees < 0:
                            print("Please enter a positive integer or 0")
                            numEmployees = int(input("\nHow many employees were on the clock: "))
                            
                        gottenNumEmp = True
                        
                    except:
                        
                        print("ERROR: Number of employees must be integers (whole numbers).")
                        continue
                
                # Creates loop based on number of employees on the clock that day
                while empNum <= numEmployees:
                    
                    empName = str(input("\nEnter employee #" + str(empNum) + "'s name: "))
                    empList.append(empName)
                    empNum += 1
                
                # Splits employee lists as strings joined by commas and display with proper capitalization    
                nameString = ", ".join(empList)
                nameString = nameString.title()
                
                # Enters users input into workbook saves and exits infinite loop    
                if day == "thursday":                
                        
                    weekSheet["B2"] = storedMileage
                    weekSheet["B3"] = mileage
                    weekSheet["B5"] = gasCost
                    weekSheet["B6"] = numStops
                    weekSheet["B7"] = nameString
                    
                    numbersBook.save(logLocation)
                    
                    file = open(mileageLocation, "w")
                    file.write(str(mileage))
                    file.close()
                    break
                        
                if day == "friday":
                    weekSheet["C2"] = storedMileage
                    weekSheet["C3"] = mileage
                    weekSheet["C5"] = gasCost
                    weekSheet["C6"] = numStops
                    weekSheet["C7"] = nameString
                    
                    numbersBook.save(logLocation)
                    
                    file = open(mileageLocation, "w")
                    file.write(str(mileage))
                    file.close()
                    break
                        
                if day == "saturday":
                    weekSheet["D2"] = storedMileage
                    weekSheet["D3"] = mileage
                    weekSheet["D5"] = gasCost
                    weekSheet["D6"] = numStops
                    weekSheet["D7"] = nameString
                    
                    numbersBook.save(logLocation)
                    
                    file = open(mileageLocation, "w")
                    file.write(str(mileage))
                    file.close()
                    break
                        
                if day == "sunday":
                    weekSheet["E2"] = storedMileage
                    weekSheet["E3"] = mileage
                    weekSheet["E5"] = gasCost
                    weekSheet["E6"] = numStops
                    weekSheet["E7"] = nameString
                    
                    numbersBook.save(logLocation)
                    
                    file = open(mileageLocation, "w")
                    file.write(str(mileage))
                    file.close()
                    break
                        
                if day == "monday":
                                
                    weekSheet["F2"] = storedMileage
                    weekSheet["F3"] = mileage
                    weekSheet["F5"] = gasCost
                    weekSheet["F6"] = numStops
                    weekSheet["F7"] = nameString
                    
                    numbersBook.save(logLocation)
                    
                    file = open(mileageLocation, "w")
                    file.write(str(mileage))
                    file.close()
                    break
                        
                if day == "tuesday":
                    
                    weekSheet["G2"] = storedMileage
                    weekSheet["G3"] = mileage
                    weekSheet["G5"] = gasCost
                    weekSheet["G6"] = numStops
                    weekSheet["G7"] = nameString
                    
                    numbersBook.save(logLocation)
                    
                    file = open(mileageLocation, "w")
                    file.write(str(mileage))
                    file.close()
                    break
                        
                if day == "wednesday":
                    weekSheet["H2"] = storedMileage
                    weekSheet["H3"] = mileage
                    weekSheet["H5"] = gasCost
                    weekSheet["H6"] = numStops
                    weekSheet["H7"] = nameString
                    
                    numbersBook.save(logLocation)
                    
                    file = open(mileageLocation, "w")
                    file.write(str(mileage))
                    file.close()
                    break
            
        # resets infinite loop if original inputs are unrecognizable
                    
            else:
                
                print("enter valid day. (Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Sunday)\n")   
                continue
                
        else:
            
            print("Period week must be in range (1, 2, 3, 4)\n")
            continue
        