# admin classes
class login:

    def __init__(self, email_address, password, type, last_login_date):
        self.email_address = email_address
        self.password = password
        self.type = type
        self.last_login_date = last_login_date

    def describe(self):
        print("Enter Login")


class log:
    def __init__(self, date, event_type, message):
        self.date = date
        self.event_type = event_type
        self.message = message

    def describe(self):
        print()


class admin(login):
    def username(self):
        print("")

    def password(self):
        print("")


class manager(login):
    def username(self):
        print("")

    def password(self):
        print("")


class employee(login):
    def username(self):
        print("")

    def password(self):
        print("")


class vendor(login):
    def username(self):
        print("")

    def password(self):
        print("")
