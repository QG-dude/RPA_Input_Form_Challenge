import os
import urllib
import pandas as pd


class Excelfile :
    def __init__(self, fileName, folderPath):
        self.fileName = fileName
        self.filePath = os.path.join(folderPath, fileName)
        self.alreadyInPlace = os.path.exists(self.filePath)
    def print_path(self):
        print(f"{self.filePath}")
    def delete(self):
        print(self.filePath)
        if self.alreadyInPlace:
            os.remove(self.filePath)
            self.alreadyInPlace = False
            print(f"{self.fileName} has been removed.")
        else:
            print(f"{self.fileName} is not found in the folder, so it cannot be removed.")
    def download(self):
        if not self.alreadyInPlace:
            print(f"Downloading of {self.fileName}...")
            fileURL = "https://rpachallenge.com/assets/downloadFiles/challenge.xlsx"
            destination = self.filePath
            # Create the request with User-Agent and open the file
            req = urllib.request.Request(fileURL, headers={"User-Agent": "Mozilla/5.0"})
            response = urllib.request.urlopen(req)
            # Save the file in the destination folder
            with open(destination, "wb") as file:
                file.write(response.read())
            # Update the alreadyInPlace variable
            self.alreadyInPlace = os.path.exists(self.filePath)
            print(f"Download completed : {destination}")
        else:
            print(f"File is already present in the folder : {self.filePath}")
    def readFile(self):
        df=pd.read_excel(self.filePath)
        return df

class Information:
    def __init__(self, firstName, lastName, companyName, roleInCompany, address, email, phone):
        self.firstName = firstName
        self.lastName = lastName
        self.companyName = companyName
        self.roleInCompany = roleInCompany
        self.address = address
        self.email = email
        self.phone = str(phone)

class Form :
    def __init__(self):
        self.selectors = {
            "firstName":'//*[@ng-reflect-name="labelFirstName"]',
            "lastName":'//*[@ng-reflect-name="labelLastName"]',
            "companyName":'//*[@ng-reflect-name="labelCompanyName"]',
            "roleInCompany":'//*[@ng-reflect-name="labelRole"]',
            "address":'//*[@ng-reflect-name="labelAddress"]',
            "email":'//*[@ng-reflect-name="labelEmail"]',
            "phone":'//*[@ng-reflect-name="labelPhone"]',
        }
