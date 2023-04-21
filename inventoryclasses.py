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
            # items.append(
            # f"{item.item_type}, {item.item_id}, {item.item_name}, {item.quantity}")
        # print(items)
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
                    # cart.append(item)
            try:
                wb.save(self.file_path)
            except:
                print("\nInvalid input")
            finally:
                show_inventory_menu()
        return cart

    # def restock(self):
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
            item.price = new_price
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
        threshold = 30.0

        for item in self.items:
            if item.quantity < threshold:
                low_inventory_items.append(item)
        # for i in low_inventory_items:

        if len(low_inventory_items) == 0:
            low_inventory_items = None
            print("No items are low on inventory")
        elif len(low_inventory_items) > 0:
            return low_inventory_items

    def place_order(self, orders):
        items = self.get_inventory()
        updated_items = []
        total_price = 0
        counter = 0
        try:
            for order in orders:
                item = next((x for x in items if x.item_id ==
                            order.item_id), None)
                if item is None:
                    continue
                item.quantity = 50
                updated_items.append(item)
                total_price += Order.calculate_total_price(InventoryManager.load_items(
                    get_file_location()), InventoryManager.get_low_inventory(get_file_location()))
        except:
            return

        order_id = self.generate_order_id()
        log_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        order_log = f"Order ID: {order_id}, Total Price: {total_price}, Time: {log_time}"
        self.log_order(order_log)
        for item in updated_items:

            counter += 1

    def generate_order_id(self):
        return str(random.randint(1000, 9999))

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
                total += item.price * item.quantity
            return total


class OrderManager:
    def __init__(self, inventory_manager):
        self.inventory_manager = inventory_manager

    def place_order(self, orders):
        items = self.inventory_manager.get_inventory()
        updated_items = []
        total_price = 0
        for order in orders:
            item = next((x for x in items if x.item_id ==
                         order.item.item_id), None)
            if item is None:
                continue
            item.quantity = 50
            updated_items.append(item)
            total_price += order.calculate_total_price(InventoryManager.load_items(
                get_file_location()), InventoryManager.get_low_inventory(get_file_location()))

        order_id = self.generate_order_id()
        # log_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # order_log = f"Order ID: {order_id}, Total Price: {total_price}, Time: {log_time}"
        # self.log_order(order_log)

        # self.inventory_manager.update_items(updated_items)

    def generate_order_id(self):
        return str(random.randint(1000, 9999))

    def log_order(self, order_log):
        try:
            with open("order.log", "a") as f:
                f.write(order_log + "\n")
        except Exception as e:
            print(f"Error writing to order log: {e}")


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
            return new_price
        except ValueError:
            print("\nYou must enter numbers")


def show_inventory_menu():
    choice = str(input("\n1. Update item price\n2. Check Inventory\n4. Get Low Inventory\n5. Calculate Total for Low Inventory Order\n6. Place Order\n7. Generate Order ID\n8. Log Order\n9. Exit\n"))
    if choice == "1":
        x = correct_price()
        InventoryManager.set_price(get_file_location(), InventoryManager.get_cart(
            get_file_location(), x), x)
    elif choice == "2":
        print(InventoryManager.get_inventory(get_file_location()))
        show_inventory_menu()
    elif choice == "4":
        print("\n")
        if InventoryManager.get_low_inventory(get_file_location()) is None:
            show_inventory_menu()
        else:
            print(InventoryManager.get_low_inventory(get_file_location()))
            show_inventory_menu()
    elif choice == "5":
        Order.calculate_total_price(
            InventoryManager.load_items(get_file_location()), InventoryManager.get_low_inventory(get_file_location()))
    elif choice == "6":
        InventoryManager.place_order(get_file_location(),
                                     InventoryManager.get_inventory(get_file_location()))
    elif choice == "7":
        OrderManager.generate_order_id(
            get_file_location())  # Location of the excel file
    elif choice == "8":
        OrderManager.log_order()
    elif choice == "9":
        print("\nHave a great day! Closing Now...")
        exit()
    else:
        print("Invalid")
        show_inventory_menu()
