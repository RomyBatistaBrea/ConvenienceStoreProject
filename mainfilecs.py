# Main file
import adminclasses
import financeclasses2
import hrclasses
import inventoryclasses


business = financeclasses2.Finance(
    1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19)
employee1 = hrclasses.Employee(
    "Jaden", "Batista", 21, "email@gmail.com", "cashier", True, 15.50, "12/12/2022")
# invManager = inventoryclasses.InventoryManager(
# '/Users/romyb/Downloads/Inventory.xlsx')
# ordermanager = inventoryclasses.OrderManager(1)
# ^ this was to test the inventory menu
answer = str(input("1. Customer\n2. Other\n3. Exit\n"))

if answer == '1':
    print("You're in customer controls")
elif answer == '2':
    # Login Function to check their job position
    check = str(input("\n1. Admin\n2. Finance\n3. HR\n4. Inventory\n"))
    if check == '1':  # Admin
        print("You're in admin controls")
    elif check == '2':  # Finance
        business.show_finance_menu()
    elif check == '3':  # HR
        employee1.show_hr_menu()
    elif check == '4':  # Inventory
        inventoryclasses.show_inventory_menu()
elif answer == '3':
    print("Thanks for using our system")
else:
    print("Invalid")
