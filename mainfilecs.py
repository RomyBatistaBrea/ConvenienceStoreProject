# Main file
import pyautogui
import adminclasses


file_path = 'C:/Users/romyb/Downloads/Inventory.xlsx'


def start_check():
    try:
        pyautogui.getWindowsWithTitle(
            'Excel')[0].activate()
        pyautogui.hotkey('alt', 'f4')
        pyautogui.sleep(1)
    except:
        pass

    answer = str(
        input("1. Customer\n2. Associate\n3. Exit\n----------------------------\n"))
    if answer == '1':
        adminclasses.customer()
    elif answer == '2':
        # Login
        # Function to check their job position
        adminclasses.login()
    elif answer == '3':
        print("Closing...")
        exit()
    else:
        print("\nInvalid\n")
        start_check()


start_check()
