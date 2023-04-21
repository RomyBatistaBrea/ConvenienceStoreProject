import os


class Employee:
    # Employment_Status, Task):
    def __init__(self, First_Name, Last_Name, Age, Email, Position, Scheduled, Pay_Rate, Hired_Date):
        self.First_Name = First_Name
        self.Last_Name = Last_Name
        self.Age = Age
        self.Email = Email
        self.Position = Position
        self.Scheduled = bool(Scheduled)
        self.Pay_Rate = Pay_Rate
        self.Hired_Date = Hired_Date
        # self.Fired_Date = Fired_Date
        # self.Employment_Status = bool(Employment_Status)
        # self.Task = Task

    def hire(self):
        print("You are in 'hire' method")
        # self.Employment_Status = True
        # print("Employee is hired.")

    def fire(self):
        print("You are in 'fire' method")
        # firing_reason = input("Reason for Firing?")
        # self.Employment_Status = False
        # print("Employee has been fired for" + firing_reason)

    def add_to_schedule(self):
        print("You are in 'add' to schedule methods")
        # self.Scheduled = True
        # print("Employee is now on the schedule.")

    def view_scheudle(self):
        print("You are in 'view' schedle method")

    def edit_schedule(self):
        print("You are in 'edit' schedule method")

    def remove_from_schedule(self):
        print("You are in 'remove' from schedule method")
        # self.Scheduled = False
        # print("Employee is now off the schedule.")

    def view_trainings(self):
        print("You are in 'view' training method")

    def record_completion(self):
        print("You are in 'record' training completion method")

    def show_hr_menu(self):
        choice = str(input("\n1. Hire\n2. Fire\n3. Add to Schedule\n4. View Schedule\n5. Edit Schedule\n6. Remove from Schedule\n7. View Trainings\n8. Record Training Completion\n9. Exit\n"))
        if choice == "1":
            self.hire()
        elif choice == "2":
            self.fire()
        elif choice == "3":
            self.add_to_schedule()
        elif choice == "4":
            self.view_scheudle()
        elif choice == "5":
            self.edit_schedule()
        elif choice == "6":
            self.remove_from_schedule()
        elif choice == "7":
            self.view_trainings()
        elif choice == "8":
            self.record_completion()
        elif choice == "9":
            return
        else:
            print("Invalid choice")
            self.show_hr_menu()
