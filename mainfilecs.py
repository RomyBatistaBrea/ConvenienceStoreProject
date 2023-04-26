# Main file
import adminclasses


def start_check():
    answer = str(input("1. Customer\n2. Associate\n3. Exit\n"))
    if answer == '1':
        print("You're in customer controls")
    elif answer == '2':
        # Login Function to check their job position
        adminclasses.login()
    elif answer == '3':
        print("Closing...")
        exit()
    else:
        print("\nInvalid\n")
        start_check()


start_check()
