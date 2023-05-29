import os
import win32com.client as win32
import sys


# How to use
# python3 .\spam.py C:\Users\Rory\Desktop\test_NO.pst
# Replace "C:\path\to\pst\file.pst" with the actual path to your PST file

pst_file_path = sys.argv[1]

print("Loading PST into OUTLOOK")
# Check if the PST file exists
if not os.path.exists(pst_file_path):
    print("PST file does not exist.")
    exit()

outlook = win32.Dispatch("Outlook.Application")

try:
    # Get the Outlook namespace
    namespace = outlook.GetNamespace("MAPI")

    # Open the PST file
    pst = namespace.AddStore(pst_file_path)
except Exception as e:
    if "The Outlook data file (.pst) failed to load for this session" in str(e):
        print("")
        print("An error occurred: The Outlook data file (.pst) failed to load for this session.")
        print("Please follow these steps:")
        print('1. Launch Outlook on your computer.')
        print('2. Go to the "File" tab in the Outlook menu.')
        print('3. Select "Open & Export" from the options on the left.')
        print('4. Choose "Open Outlook Data File" from the list.')
        print('5. Browse and select the password-protected PST file you want to load into Outlook.')
        print('6. Click "Open" to start the loading process.')
        print('7. A dialog box will appear asking you to enter the password for the PST file.')
        print('8. Enter the correct password for the PST file and click "OK".')
        print('9. If the password is correct, Outlook will load the PST file, and it will appear as a separate mailbox or folder in the Outlook navigation pane.')
        print("")
        exit()

    else:
        print(f"An error occurred: {str(e)}")
        exit()


finally:
    # Close the PST file
    if 'pst' in locals() and pst is not None:
        pst.Close()

    # Release the COM objects
    if 'namespace' in locals():
        namespace = None
    if 'outlook' in locals():
        outlook = None
