
from os import path,  scandir, makedirs, remove
from sys import exit
from shutil import copy
import win32com.client as win32
import subprocess
import time
#python -m PyInstaller -F -c main.py
# the folders made by the program do not seem to load in sharepoint webbrowser view.

class FATPunchlistExtractor(object):
    def __init__(self):
        self.punchlist_path = r"\\trianglepackagemachines.sharepoint.com@SSL\DavWWWRoot\service\techs\punchlists"
        self.testing_path = r"\\trianglepackagemachines.sharepoint.com@SSL\DavWWWRoot\service\techs\testing"
        self.fat_path = r"\\trianglepackagemachines.sharepoint.com@SSL\DavWWWRoot\service\customer\Shared Documents" # this is the customer punclist path
        self.using_alt_path = False
        self.instructions = "In the Triangle punchlist, every numbered row you'd like to show up on the customer punchlist\n " \
                            "should be highlighted red via the 'BAD' cell style, or one of the 2 standard reds.\n" \
                            "Highlight ONLY the numbers, column 1, red. Every unhighlighted row gets deleted."


    def main(self):
        rerun = True
        if not self.connection_check_SP():
            input("Exiting. Press Enter...")
            quit()
        print(self.instructions)
        punchlist_path = self.serial_number_entry()
        self.copy_FAT_items(punchlist_path)


    def connection_check_SP(self):
        connected = path.exists(r"\\trianglepackagemachines.sharepoint.com@SSL\DavWWWRoot\service")
        if not connected:
            print("No SharePoint Connection Exists")
        return connected

    def serial_number_entry(self):
        xl_path = None
        while xl_path is None:
            serial_no = input("Enter a Serial Number or enter 'quit' to exit: ")
            if serial_no == "quit":
                exit()
            if len(serial_no) == 6:
                xl_path = self.find_punchlist(serial_no, self.punchlist_path)
                if xl_path == None:
                    xl_path = self.find_punchlist(serial_no, self.testing_path)
                    self.using_alt_path = True
                if xl_path == None:
                    print("No punchlists found for serial number %s" % serial_no)
            else:
                print("That's not a valid Serial Number.")
        print("Punchlist found at:\n" + xl_path)
        return xl_path


    def find_punchlist(self, serial_no, folder):
        paths = []
        for file in scandir(folder):
            if serial_no in file.name:
                paths.append(file)
        if len(paths) > 1:
            file_path = paths[0]
            print("Multiples Files found, utility will use most recent.")
            for file in paths[0::]:
                if path.getmtime(file_path) < path.getmtime(file):
                    file_path = path
            return file_path.path
        elif len(paths) == 1:
            file_path = paths[0]
            return file_path.path
        else:
            print("Punchlist for serial number %s not found in %s" % (serial_no, folder))
            return None
        # using the entered serial number, look through all the punchlists in sharepoint until we find a match
        # if there is more than one match, notify the user and use the most upto date one, according to user input
        # if no punchlist is found, use teh old folder locations, but notify the user
        # this function will return the true path of the punchlist found

    def init_customer_punchlist(self, true_path):
        # this function returns the name of the customer FAT punchlist, aka the true path for use when saving
        # returns the worksheet to copy the rows onto
        xl = win32.gencache.EnsureDispatch('Excel.Application')
        xl.Visible = True
        xl.DisplayAlerts = False
        wb = xl.Workbooks.Open(true_path)
        for ws in wb.Worksheets:
            if ws.Name != "Testing":
                ws.Delete()
        wb.SaveAs(self.fat_path + "\\" + true_path.split("\\")[-1])

    def excel_colors(self):
        xl = win32.gencache.EnsureDispatch('Excel.Application')
        xl.Visible = True
        xl.DisplayAlerts = False
        # xl.AutoRecover.Enabled = False
        wb = xl.Workbooks.Open(r"C:\Users\sotero\Documents\124551_HANOVER FOODS_Punchlist.xlsx")
        start_row = 17
        c_row = start_row  # current row

        ws = wb.Worksheets("Testing")
        while (ws.Cells(c_row, 1).Value is not None or
               ws.Cells(c_row + 1, 1).Value is not None or
               ws.Cells(c_row + 2, 1).Value is not None or
               ws.Cells(c_row + 3, 1).Value is not None):
            print(ws.Cells(c_row, 1).Interior.Color)
            print(ws.Cells(c_row, 1).Interior.Color not in [13551615.0, 192.0, 255.0])
            c_row+=1

    def copy_FAT_items(self, true_path):
        # this function will open both the punchlist and the customer punchlist,
        # assign variables to the 2 worksheets
        # run through all the rows of the test list
        # to get past the funky punchlists with merged rows, check that the number increments every row.
        # determine the number of columns to use
        # determine the active issue
        # determine the active issues to copy over,
        # notifying the user at each turn
        # then copy the rows to the FAT list.
        # there will be some gimmicky code due to the fact that the punchlists use 4 rows in per row.
        fat_full_path = self.fat_path + "\\" + true_path.split("\\")[-1]
        #fat_full_path = r"C:\Users\sotero\Documents\Misc" + "\\" + true_path.split("\\")[-1]
        while path.exists(fat_full_path):
            try:
                remove(fat_full_path)
                print("Old file Deleted")
            except PermissionError:
                print("Original File is open, please close it to continue.")
                old_file_stop = input("Press Enter to continue, or type 'quit' then enter to exit:")
                if old_file_stop in ["quit", "'quit'"]:
                    quit()

        copy(true_path, fat_full_path)
        xl = win32.gencache.EnsureDispatch('Excel.Application')
        xl.Visible = True
        xl.DisplayAlerts = False
        #xl.AutoRecover.Enabled = False
        wb = xl.Workbooks.Open(fat_full_path)

        for ws in wb.Worksheets:
            if ws.Name != "Testing":
                ws.Delete()
        ws = wb.Worksheets("Testing")
        start_row = 17
        c_row = start_row  # current row
        row_increment = 1  # used to renumber the rows after rows in between selected rows are deleted

        while (ws.Cells(c_row, 1).Value is not None or
                ws.Cells(c_row + 1, 1).Value is not None or
                ws.Cells(c_row + 2, 1).Value is not None or
                ws.Cells(c_row + 3, 1).Value is not None):
            #print(ws.Cells(c_row, 1).Interior.Color)
            if ws.Cells(c_row, 1).Interior.Color not in [13551615.0, 192.0, 255.0]:
                ws.Rows(c_row).EntireRow.Delete()
                c_row -= 1
            else:
                #ws.Cells(c_row, 1).Value = row_increment
                row_increment += 1
            c_row += 1

        wb.Save()
        wb.Close(False)
        print("Finished, check %s for customer Punchlist" % fat_full_path)
        subprocess.call('explorer ' + fat_full_path, shell=True)
        #self.folder_find(fat_full_path)

    def folder_find(self, starting_path):
        # first split the path up to find the customer name
        customer_name_unformatted = self.find_customer_name(starting_path.split("_")[1])
        customer_name = self.format_customer_name(customer_name_unformatted)
        customer_folder = None
        # search through the folder to find one with the same name
        # since the folder names are to be as short as possible, check if folder name is in the customer string
        # exclude the site asset, lists, images and all other standard sharepoint folders.
        main_folder = r"\\trianglepackagemachines.sharepoint.com@SSL\DavWWWRoot\service\customer"
        standard_folders = ["Lists", "images", "Shared Documents", "SiteAssets", "SitePages"]
        for folder in scandir(main_folder):
            if folder.name.lower() in customer_name.lower() and folder.name not in standard_folders:
                customer_folder = folder
                break
        # if a folder name is found, confirm with the user whether to save in the found location.
        if customer_folder is not None:
            print("Found folder '" + customer_folder.name + "' in the customer Sharepoint.")
            use_folder = input("Save punchlist in this folder?(y/n)")
            if use_folder == "y":
                print("Saving File to folder, original copy will not be affected")
                copy(starting_path, customer_folder.path + "\\" + starting_path.name)
                end_path = customer_folder.path + "\\" + starting_path.name
            else:
                print("Folder Rejected, shutting down.")
                time.sleep(2)
                print("jk watchu want?")
        else:
            # if no folder is found, use first word in the customer name (not if it's a form of be) to suggest folder creation
            # confirm with user is folder is to be created (capitalize the first letter)
            create_folder = input("No Folder found... Create Folder named: " + customer_name + "?(y/n)")
            if create_folder == "y":
                # I mix data types here because i cant find the os command to get the path like object to the new directory
                customer_folder = main_folder + "\\" + customer_name
                try:
                    makedirs(customer_folder)
                    print("Folder created at " + customer_folder)
                except FileExistsError:
                    print("This folder already existed, file will be copied there.")
            else:
                name_input = input("Type new name for folder an press enter(don't type \ / ; * ? '' < >):")
                customer_folder = main_folder + "\\" + name_input
                try:
                    makedirs(customer_folder)
                    print("Folder created at " + customer_folder)
                except FileExistsError:
                    print("This folder already existed, file will be copied there.")
            print("Saving File to new folder, original copy will not be affected")
            copy(starting_path, customer_folder + "\\" + starting_path)
            end_path = customer_folder + "\\" + starting_path
            subprocess.call('explorer ' + end_path, shell=True)

    def format_customer_name(self, raw_name):
        lc_name = raw_name.lower()
        cap_first_letter = lc_name[0].upper() + lc_name[1::]
        return cap_first_letter

    def find_customer_name(self, file_name):
        forms_of_be = ["a", "the"]
        names = file_name.split(" ")
        for name in names:
            if name.lower() not in forms_of_be or len(name) > 2:
                return name
        return "Really Short Customer Name"


if __name__ == "__main__":
    rs = FATPunchlistExtractor()
    while True:
        try:
            rs.main()
        except Exception as ex:
            print("Error Found: {}, contact your nerd.".format(ex))


