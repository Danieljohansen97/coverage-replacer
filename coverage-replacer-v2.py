# Made by Daniel Johansen 22. December 2023
# v2.0
import pandas as pd
import time

startTime = time.time()
endStatus = ""

## Path variables
svvExcelPath = "excel-files/SVV.xlsx"
coverageExcelPath = "excel-files/coverage-descriptions.xlsx"

## Create a dataframe per excel sheet and print error if not found
try:
    svvDF = pd.read_excel(svvExcelPath, sheet_name="Sheet1")
except:
    print(f"ERROR: Could not find: {svvExcelPath}")

try:
    coverageDF = pd.read_excel(coverageExcelPath, sheet_name="Sheet2")
except:
    print(f"ERROR: Could not find: {coverageExcelPath}")

## Print Columns from dataframes as a reference
try:
    print(f'\n=== COLS IN ROAD CHANGE OVERVIEW ===\n\n{svvDF.columns}')
    print(f'\n=== COLS IN COVERAGE DESCRIPTIONS ===\n\n{coverageDF.columns}')
except:
    print("ERROR: Check dataframe imports")
# 3984 rows before change and 3984 rows after change
# Add correct county name to svv dataframe based on old county numbers
countyNameColumn = {"Fylkenavn" : []}
for index in svvDF.index:
    currentCountyNumber = svvDF["Fylke"][index]
    match currentCountyNumber:
        case 1 | 2 | 6: # Viken
            countyNameColumn["Fylkenavn"].append("Viken")
        case 3: # Oslo
            countyNameColumn["Fylkenavn"].append("Oslo")
        case 4 | 5: # Innlandet
            countyNameColumn["Fylkenavn"].append("Innlandet")
        case 7 | 8: # Vestfold og Telemark
            countyNameColumn["Fylkenavn"].append("Vestfold og Telemark")
        case 9 | 10: # Agder
            countyNameColumn["Fylkenavn"].append("Agder")
        case 11: # Rogaland
            countyNameColumn["Fylkenavn"].append("Rogaland")
        case 12 | 13 | 14: # Vestland
            countyNameColumn["Fylkenavn"].append("Vestland")
        case 15: # Møre og Romsdal
            countyNameColumn["Fylkenavn"].append("Møre og Romsdal")
        case 16 | 17: # Trøndelag
            countyNameColumn["Fylkenavn"].append("Trøndelag")
        case 18: # Nordland
            countyNameColumn["Fylkenavn"].append("Nordland")
        case 19 | 20: # Troms og Finnmark
            countyNameColumn["Fylkenavn"].append("Troms og Finnmark")
# Insert the countyNameColumn into svvDF
try:
    svvDF.insert(1, "Fylkenavn", countyNameColumn["Fylkenavn"])
except:
    print("ERROR: Could not add the new column")
# # Generate new svv excelSheet based on svvDF if needed
# # Uncomment the commented lines to generate
# try:
#     svvDF.to_excel("svv-test.xlsx")
# except:
#     print("ERROR: Could not generate excel file")

# Result statistics
changeResult = {
    "C-Yes": 0, # svv and coverage county matches
    "C-No": 0, # svv and coverage county no match
    "Fv-Yes": 0, # Fylkesveg found
    "Rv-Yes": 0, # Riksveg found
    "Skip": 0, # Nothing found, skip to next
    "Error": 0, # Errors, and unhandled values
    "index": 0, # Index of coverageDF
}
# Road combinations to scour the coverage document based on
# Both arrays should be the same length, if not the program crashes
roadCombinationsF = [
    "FV",
    " FV",
    "FV ",
    " FV ", 
    "Fv", 
    " Fv",
    "Fv ",
    " Fv ",
    "fv",
    " fv", 
    "fv ",
    " fv ",
    "Fylkesvei",
    " Fylkesvei",
    "Fylkesvei ",
    " Fylkesvei ",
    "fylkesvei",
    " fylkesvei",
    "fylkesvei ",
    " fylkesvei ",
    "Fylkesveg",
    " Fylkesveg",
    "Fylkesveg ",
    " Fylkesveg ",
    ]
roadCombinationsR = [
    "RV",
    " RV", 
    "RV ",
    " RV ",
    "Rv",
    " Rv", 
    "Rv ",
    " Rv ",
    "rv",
    " rv",
    "rv ",
    " rv ", 
    "Riksvei",
    " Riksvei",
    "Riksvei ",
    " Riksvei ",
    "riksvei",
    " riksvei",
    "riksvei ",
    " riksvei ",
    "Riksveg",
    " Riksveg",
    "Riksveg ",
    " Riksveg ",
    ]

# Block of code that does the magic
for index in coverageDF.index: # Loop through coverageDF
    changeResult["index"] = index
    coverageDescription = coverageDF["Message__C"][index]
    currentCounty = coverageDF["Fylke"][index]
    for index2 in svvDF.index:
        svvCounty = svvDF["Fylkenavn"][index2]
        svvRoad = svvDF["Veg-kategori"][index2]
        if (svvCounty == currentCounty):
            changeResult["C-Yes"] += 1
            for index3 in range(24): # Remember to change range(value) to length of roadCombinations assuming R and F are the same length         
                match svvRoad:
                    case "F":
                        # Look for fylkesveg string in coverage description
                        searchString = f"{roadCombinationsF[index3]}{svvDF['Gammelt vegnummer'][index2]}"
                        # Try to replace all occurences of searchString in coverageDescription
                        try:
                            if(coverageDescription.find(searchString) != -1):
                                changeResult["Fv-Yes"] += 1
                                newDescription = coverageDescription.replace(searchString, f"Fv{svvDF['Nytt vegnummer'][index2]}")
                                coverageDF.loc[index, ["Message__C"]] = [newDescription]
                                print(changeResult, f"[{round((index / len(coverageDF))*100, 2)}%]")
                            else:
                                changeResult["Skip"] += 1
                                print(changeResult, f"[{round((index / len(coverageDF))*100, 2)}%]")
                        except:
                            changeResult["Error"] += 1
                    case "R":
                        searchString = f"{roadCombinationsR[index3]}{svvDF['Gammelt vegnummer'][index2]}"
                        # Try to replace all occurences of searchString in coverageDescription
                        try:
                            if(coverageDescription.find(searchString) != -1):
                                changeResult["Rv-Yes"] += 1
                                newDescription = coverageDescription.replace(searchString, f"Rv{svvDF['Nytt vegnummer'][index2]}")
                                coverageDF.loc[index, ["Message__C"]] = [newDescription]
                                print(changeResult, f"[{round((index / len(coverageDF))*100, 2)}%]")
                            else:
                                changeResult["Skip"] += 1
                                print(changeResult, f"[{round((index / len(coverageDF))*100, 2)}%]")
                        except:
                            changeResult["Error"] += 1
                    case _:
                        changeResult["Error"] += 1
        else:
            changeResult["C-No"] += 1
            print(changeResult, f"[{round((index / len(coverageDF))*100, 2)}%]")

# Write the new coverage descriptions to an excel file
try:
    coverageDF.to_excel("new-coverage-descriptions.xlsx")
    endStatus = "SUCCESS, new coverage descriptions has been written"
except:
    endStatus = "ERROR: New coverage descriptions was not written to file"

endTime = time.time()
elapsedTime = (endTime - startTime) / 60
print(f"\n\n=== STATUS SUMMARY ===")
print('Status:', endStatus)
print('Execution time:', elapsedTime, 'minutes')
print(changeResult)

# TODO: Replace the hard coded range(24), should be automatically detected based on roadCombinations array length.
# TODO: Optimize and rework roadCombinations, the current solution is a hazzle to update.
# TODO: Add logic to handle cases where several road numbers follow the Fv/Rv: For example: Fv435, 506, 35, 67. I have seen this appear in the document.
# TODO: Refactor the massive main block of code and extract into its own functions for readability.