##############################################################################
# Phone Book
# Allows you to add, subtract, and replace contact information, similar to a 
# contact list in your phone
##############################################################################

from openpyxl import load_workbook
from openpyxl import Workbook
import pandas as pd 

#Create A New Phone Book
def Create_Book():
    wb = Workbook()
    destination_filename = "PhoneBook.xlsx"
    ws1 = wb.active
    ws1.title = "Contact Information"
    
    ws1["A1"] = "Name" 
    ws1["B1"] = "Cell Number"
    ws1["C1"] = "Home Number"
    ws1["D1"] = "Work Number"
    ws1["E1"] = "Email"
    ws1["F1"] = "Relationship"
    
    wb.save(filename = destination_filename)
    
    return(wb)

#Open Existing Phone Book 
def Open_Book():
    print("What is the file name of your current book?")
    destination_filename = input("Answer:")
    
    while True:
        try: 
            wb = load_workbook(destination_filename)
            break
        
        except:
            print("I'm sorry but there was an issue opening the file you specified. Please type it again.")
            destination_filename = input("Answer:")
            continue
        
    return(wb)
    
    

#Excel Sheet to DataFrame 
def Sheet_To_Frame(): 
    pass

#Add to the DataFrame
def Add_Contact():
    pass

#Subtract from the DataFrame 
def Remove_Contact():
    pass

#Replace Information from the DataFrame
def Edit_Contact():
    pass

#View All Contacts 
def View_Contacts(): 
    pass

#Search for Contact
def Search_Contact():
    pass

#Save
def Save_Book():
    pass

#Ask if they are a new user then make a new excel file 
#If they aren't a new user then let them see on an existing file (will have to \ask for the file)

print("Are you a new user? Please answer yes or no.")
while True:
    new_user_q = input("Answer:")
    new_user_q = new_user_q.lower()
    
    while True: 
        if new_user_q == "yes": 
            wb = Create_Book()
            break
            
        elif new_user_q == "no":
            wb = Open_Book()
            break
             
        else: 
            print("Sorry that answer was incorrect. Please try again")
            continue
        

    # Category Options 
    # A : Add New Contact 
    # R : Remove Contact 
    # E : Edit Existing Contact 
    # V : View Contacts 
    # S : Search For Contact 


