import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

def get_subfolders(selected_folder, folder, selected_folder_names, indent=0):
    subfolders = folder.Folders
    if folder.Name in selected_folder_names:
        selected_folder_names.remove(folder.Name)
        # Do something with the selected folder and its subfolders
        if folder == selected_folder:
            print("Selected folder:", folder.Name)
            # ...
            if len(selected_folder_names) == 0:
                return
            
    if len(subfolders) > 0:
        for subfolder in subfolders:
            print(" " * indent + subfolder.Name)
            get_subfolders(selected_folder, subfolder, selected_folder_names.copy(), indent+4)

def print_all_folders():
    top_level_folders = outlook.Folders
    selected_folders = []
    while True:
        if len(selected_folders) == 0:
            print("Available folders:")
            for i, folder in enumerate(top_level_folders):
                print(f"{i+1}. {folder.Name}")
        else:
            print("Available subfolders:")
            for i, folder in enumerate(selected_folders[-1].Folders):
                print(f"{i+1}. {folder.Name}")
        selected_folder_numbers = input("Enter the number(s), a to save or q to exit: ")
        if selected_folder_numbers.lower() == "a":
            global subfolder_location
            subfolder_location = folder.FolderPath
            break
        elif selected_folder_numbers.lower() == "q":
            break
        
        selected_folder_numbers = selected_folder_numbers.split(",")
        if len(selected_folder_numbers) > 3:
            print("Cannot select more than 3 folders.") # keep this at 3 to stop the other scripts working at this level
            continue
        selected_folders_copy = selected_folders.copy()
        for folder_number in selected_folder_numbers:
            try:
                folder_number = int(folder_number)
            except ValueError:
                print("Invalid input:", folder_number)
                selected_folders = selected_folders_copy
                break
            if len(selected_folders) == 0:
                if folder_number < 1 or folder_number > len(top_level_folders):
                    print("Invalid folder number:", folder_number)
                    selected_folders = selected_folders_copy
                    break
                selected_folders.append(top_level_folders[folder_number-1])
            else:
                subfolders = selected_folders[-1].Folders
                if folder_number < 1 or folder_number > len(subfolders):
                    print("Invalid folder number:", folder_number)
                    selected_folders = selected_folders_copy
                    break
                selected_folders.append(subfolders[folder_number-1])
        else:
            selected_folder = selected_folders[-1]
            selected_folder_names = [selected_folder.Name]
            get_subfolders(selected_folder, selected_folder, selected_folder_names, indent=4)

print_all_folders()


## this is to be used for later scripts as it provides a global variable to use

print(subfolder_location)
