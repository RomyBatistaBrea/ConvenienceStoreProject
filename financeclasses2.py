import openpyxl
from inventoryclasses import InventoryManager, get_file_location


class Finance:
    def __init__(self, insurance, product, price, benefits, orders, utilities, taxes, sales, remaining,
                 hours, maintenance, rate, gas_tax, state_tax, federal_tax, gas_used, unbilled_revenue, primary_cost, bills):
        self.insurance = insurance
        self.product = product
        self.price = price
        self.benefits = benefits
        self.orders = orders
        self.utilities = utilities
        self.taxes = taxes
        self.sales = sales
        self.remaining = remaining
        self.hours = hours
        self.maintenance = maintenance
        self.expenses = rate
        self.gas_tax = gas_tax
        self.state_tax = state_tax
        self.federal_tax = federal_tax
        self.gas_used = gas_used
        self.unbilled_revenue = unbilled_revenue
        self.primary_cost = primary_cost
        self.bills = bills

    def check_insurance_expense(self):
        # return self.insurance * self.expenses
        print("Put the insurance Plan You Want to Check")

    def calculate_profit(self):
        # profit = self.sales - (self.benefits + self.orders + self.utilities + self.taxes)
        # return profit
        print("Calculate the profit")

    def show_sales(self):
        # return self.solds, self.remaining
        print("Checks the Sales")

    def compute_maintenance_cost(self):
        # total_cost = self.hours * self.maintenance * self.price + self.taxes
        # return total_cost
        print("Checks the bills")

    def change_employee_hourly_rate(self, new_expense):
        # self.expenses = new_expense
        print(f"New expense: {new_expense}")

    def calculate_taxes(self):
        # state_tax = self.state_tax * self.sales
        # federal_tax = self.federal_tax * self.unbilled_revenue
        # gas_tax = self.gas_tax * self.gas_used
        # return gas_tax, state_tax, federal_tax
        print("Which Type of Taxes do you want to check")

    def calculate_vendor_profit_loss(self):
        # profit_loss = self.sales - (self.primary_cost + self.bills + self.taxes)
        # return profit_loss
        print("Here's How Much Profit You Made:-------")

    def show_finance_menu(self):
        choice = str(input("\n1. Check Insurance Expense\n2. Set Price\n3. Calculate Profit\n4. Show Sales\n5. Compute Maintenance Cost\n6. Change Employee Hourly Rate\n7. Calculate Taxes\n8. Calculate Vendor Profit/Loss\n9. Exit\n"))
        if choice == "1":
            self.check_insurance_expense()
        elif choice == "2":
            new_price = float(input("Enter the new price: "))
            self.set_price(InventoryManager.get_cart(
                get_file_location()), new_price)
        elif choice == "3":
            self.calculate_profit()
        elif choice == "4":
            self.show_sales()
        elif choice == "5":
            self.compute_maintenance_cost()
        elif choice == "6":
            new_expense = float(input("Enter the new expense: "))
            self.change_employee_hourly_rate(new_expense)
        elif choice == "7":
            self.calculate_taxes()
        elif choice == "8":
            self.calculate_vendor_profit_loss()
        elif choice == "9":
            return
        else:
            print("Invalid Choice")
            self.show_finance_menu(self)
