import functions

chatName = input("Enter chat name/ group name: ")
Date = input("Enter cutoff date (format: dd/mm/yyyy): ")
while functions.is_valid_datetime(Date) == False:
    print("Please enter correct date and format")
    Date = input("Enter cutoff date (format: dd/mm/yyyy): ")
check = input("Headless?(1 for yes, 0 for no): ")
if int(check) == 1:
    check = True
else:
    check = False


messages = functions.getMessages1(chatName, Date, check)
response = int(input("Enter 0 for direct storage, 1 for storage in temp sheet: "))
workBook = "QJS Daily Sales Report  30-Mar-2023.xlsx"
if int(response) == 1:
    sheetName = input("Enter sheet name: ")
    if functions.storeDate(messages, response, workBook, sheetName) == 1:
        print("done!")
    else:
        print("Storing procedure failed")
else:
    if functions.storeDate(messages, response, workBook, 0) == 1:
        print("done!")
    else:
        print("Storing procedure failed")
