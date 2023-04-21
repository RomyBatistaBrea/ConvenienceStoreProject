# datetime to log orders
# random to generate order ID's for customers
# You have to install "openpyxl" to to run code.
# open powershell for windows type in "pip install openpyxl"
# restart computer
# Finance has to add tab for price and update class price and tax calcualtions

import datetime
import random
import openpyxl


class Item:
    def __init__(self,  item_type, item_id, item_name, quantity, price):
        self.item_type = item_type
        self.item_id = item_id
        self.item_name = item_name
        self.quantity = quantity
        self.price = price


class InventoryManager:
    def __init__(self, file_path):
        self.file_path = file_path
        self.items = self.load_items()

    def load_items(self):
        try:
            wb = openpyxl.load_workbook(self.file_path)
        except FileNotFoundError:
            print(f"Error: Could not find file {self.file_path}")
            return []

        sheet = wb.active
        items = []
        for row in sheet.iter_rows(values_only=True):
            item = Item(item_id=row[1], item_name=row[2],
                        item_type=row[0], quantity=row[3],
                        price=row[4])
            if item.quantity == "Quantity":
                continue
            else:
                items.append(item)  # Appends items as objects
        return items

    def update_items(self, items):
        max_item_capacity = 0
        try:
            wb = openpyxl.load_workbook(self.file_path)
        except FileNotFoundError:
            print(f"Error: Could not find file {self.file_path}")
            return

        # sheet = wb.active
        # sheet.delete_rows(2, sheet.max_row)  # clear existing items
        # for item in items:
        #     sheet.append([item.item_id, item.item_name,
        #                  item.item_type, max_item_capacity, item.price])
        # wb.save(self.file_path)
        sheet = wb.active
        ws = sheet
        counter = 0
        for row in sheet.iter_rows(values_only=True):
            if str(row[1]) != 50:
                ws.cell(row=(counter+1),
                        column=5).value = max_item_capacity
                wb.save(self.file_path)

    def get_cart(self, new_price):
        option = ""
        cart = []
        x = True

        print('''Enter the id of the item: ''')
        while x:
            option = input("Enter item: ")
            for item in self.load_items():
                try:
                    if int(option) == int(item.item_id):
                        try:
                            wb = openpyxl.load_workbook(self.file_path)
                        except FileNotFoundError:
                            print(
                                f"Error: Could not find file {self.file_path}")
                            return []
                        sheet = wb.active
                        ws = sheet
                        counter = 0
                        for row in sheet.iter_rows(values_only=True):
                            if option not in str(row[1]):
                                counter += 1
                            else:
                                ws.cell(row=(counter+1),
                                        column=5).value = new_price
                except Exception as e:
                    x = False
                    print(f"Error: {e}")
            try:
                wb.save(self.file_path)
            except:
                print("\nInvalid input")
            finally:
                show_inventory_menu()
        return cart

    def restock(self, updated_quantity):
        file_path = self.file_path
        try:
            wb = openpyxl.load_workbook(
                '/Users/romyb/Downloads/Inventory.xlsx')
        except FileNotFoundError:
            print(f"Error: Could not find file {file_path}")
            return

        sheet = wb.active
        sheet.delete_rows(2, sheet.max_row)  # clear existing items
        for items in updated_quantity:
            sheet.append([items[0], items[1], items[2], items[3], items[4]])
        wb.save(file_path)
    #     option = ""
    #     cart = []

    #     for item in self.load_items():
    #         try:
    #             wb = openpyxl.load_workbook(self.file_path)
    #         except FileNotFoundError:
    #             print(
    #                 f"Error: Could not find file {self.file_path}")
    #             return []
    #         sheet = wb.active
    #         ws = sheet
    #         counter = 0
    #         for row in sheet.iter_rows(values_only=True):
    #             if option not in str(row[1]):
    #                 counter += 1
    #             else:
    #                 ws.cell(row=(counter+1),
    #                         column=4).value = 50
    #     return cart

    def set_price(self, items, new_price):
        file_path = '/Users/romyb/Downloads/Inventory.xlsx'
        new_item_list = []
        for item in items:
            print([item.item_id, item.item_name, item.item_type, item.quantity])
            item.price = locale.currency(float(new_price), grouping=True)
            new_item_list.append(item)

        try:
            wb = openpyxl.load_workbook(
                '/Users/romyb/Downloads/Inventory.xlsx')
        except FileNotFoundError:
            print(f"Error: Could not find file {file_path}")
            return

        sheet = wb.active
        for item in new_item_list:
            sheet.append([item.item_id, item.item_name,
                         item.item_type, item.quantity, item.price])
        wb.save(file_path)

    def get_inventory(self):
        return self.items

    def get_low_inventory(self):
        low_inventory_items = []
        threshold = 50.0

        for item in self.items:
            if item.quantity < threshold:
                low_inventory_items.append(item)

        if len(low_inventory_items) == 0:
            low_inventory_items = None
            print("No items need to be restocked.")
        elif len(low_inventory_items) > 0:
            return low_inventory_items

    def place_order(self):
        items = self.get_inventory()
        updated_items = []
        total_price = 0
        counter = 0
      #  try:
        for item in items:
            price = str(item.price).replace(",", "")
            price = price.replace("$", "")
            price = float(price)

            if item.quantity <= 50:
                counter += 1
                total_price += price * (50 - item.quantity)

            updated_items.append([
                item.item_type, item.item_id, item.item_name, 50, locale.currency(float(price), grouping=True)])
        if counter == 0:
            print("\nAll items are stocked, none need to be ordered.")
        else:
            self.restock(updated_items)
            print(
                "All items have been restocked to 50 units and logged in the order log.\nTotal price: ${:.2f}".format(total_price))

        order_id = self.generate_order_id()
        log_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        order_log = f"Order ID: {order_id}, Total Price: ${total_price}, Time: {log_time}"
        self.log_order(order_log)

        show_inventory_menu()

    def generate_order_id(self):
        return str(random.randint(100000, 999999))

    def log_order(self, order_log):
        try:
            with open("order.log", "a") as f:
                f.write(order_log + "\n")
        except Exception as e:
            print(f"Error writing to order log: {e}")


class Order:
    def __init__(self, quantity):
        self.quantity = quantity

    def calculate_total_price(self, item_list):
        if item_list is None:
            return 0
        else:
            total = 0
            for item in item_list:
                total += item.price * (50 - item.quantity)
            return total


def get_file_location():
    return InventoryManager(
        '/Users/romyb/Downloads/Inventory.xlsx')


def correct_price():
    x = True
    while x:
        try:
            new_price = float(input("Enter the new price: "))
            while new_price < float(0.01):
                try:
                    if new_price < float(0.01):
                        print("\nPrice has to be greater than 0")
                    new_price = float(input("\nEnter the new price: "))
                except ValueError:
                    print("\nYou must enter numbers")
            x = False
            return locale.currency(float(new_price), grouping=True)
        except ValueError:
            print("\nYou must enter numbers")


def show_inventory_menu():
    choice = str(input("\n1. Update item price\n2. Check Current Inventory\n3. Get Restockable Items list\n4. Calculate Total for Low Inventory Order\n5. Order Inventory Restock\n6. Generate Order ID\n7. Log Order\n8. Exit\n"))
    if choice == "1":
        x = correct_price()
        InventoryManager.set_price(get_file_location(), InventoryManager.get_cart(
            get_file_location(), x), x)
    elif choice == "2":
        for item in InventoryManager.get_inventory(get_file_location()):
            print(
                f"{item.item_type}, {int(item.item_id)}, {item.item_name}, Quantity: {item.quantity}, Price: {item.price}")
        # print(InventoryManager.get_inventory(get_file_location()))
        show_inventory_menu()
    elif choice == "3":
        print("\n")
        if InventoryManager.get_low_inventory(get_file_location()) is None:
            show_inventory_menu()
        else:
            for item in InventoryManager.get_low_inventory(get_file_location()):
                print(f"{item.item_name}: {item.quantity} units left")
            show_inventory_menu()
    elif choice == "4":
        print("")
        x = Order.calculate_total_price(
            InventoryManager.load_items(get_file_location()), InventoryManager.get_low_inventory(get_file_location()))
        if x == 0:
            print("")
        else:
            print(f"\nTotal price for restock order: ${x}")
        show_inventory_menu()
    elif choice == "5":
        InventoryManager.place_order(get_file_location())
    elif choice == "6":
        print(InventoryManager.generate_order_id(
            get_file_location()))  # Location of the excel file
        show_inventory_menu()
    elif choice == "7":
        InventoryManager.log_order()
    elif choice == "8":
        print("\nHave a great day! Closing Now...")
        exit()
    else:
        print("Invalid")
        show_inventory_menu()
