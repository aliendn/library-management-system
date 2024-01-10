import ast
import datetime
import tkinter as tk
import uuid

import pandas as pd


def email_validator(email):
    flag = True
    if "@" in email:
        user_name = email.split("@")[0]
        domain = email.split("@")[1]
        name_domain = domain.split(".")[0]
        sub_domain = domain.split(".")[1]
        if len(user_name) > 12:
            print("your username is too long")
            flag = False
        elif not user_name[0].isalpha():
            print("your username should start with alphabetic characters")
            flag = False
        elif 2 > len(name_domain) or len(name_domain) > 5:
            print("your domain name is not correct")
            flag = False
        elif len(sub_domain) != 3:
            print("your sub domain is not correct")
            flag = False
    else:
        print("your email address is not correct")
        flag = False
    print("your email address is correct")
    return flag


def full_name_validator(name):
    return len(name) > 0


def password_validator(password):
    return len(password) >= 8


class LibraryApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.user_session = None
        self.title("My Library Managements")
        self.login_button = tk.Button(self, text="Login", command=self.show_login)
        self.login_button.grid(row=0, pady=10)
        self.register_button = tk.Button(self, text="Register", command=self.show_register)
        self.register_button.grid(row=1, pady=10)

    def redirect_main_menu_or_login(self):
        self.clear_window()
        self.login_button = tk.Button(self, text="Login", command=self.show_login)
        self.login_button.grid(row=0, pady=10)
        self.register_button = tk.Button(self, text="Register", command=self.show_register)
        self.register_button.grid(row=1, pady=10)

    def show_login(self):
        self.clear_window()
        self.login_button.grid_remove()
        self.register_button.grid_remove()

        self.label_username = tk.Label(self, text="Full Name or Email")
        self.label_username.grid(row=0, column=0)
        self.entry_username = tk.Entry(self)
        self.entry_username.grid(row=0, column=1)
        self.label_password = tk.Label(self, text="Password")
        self.label_password.grid(row=1, column=0)
        self.entry_password = tk.Entry(self, show="*")
        self.entry_password.grid(row=1, column=1)
        self.login_button = tk.Button(self, text="Login", command=self.login_user)
        self.login_button.grid(row=2, columnspan=2, pady=10)
        self.login_button = tk.Button(self, text="Login", command=self.login_user)
        self.login_button.grid(row=2, columnspan=2, pady=10)

        back_button = tk.Button(self, text="Back", command=self.redirect_main_menu_or_login)
        back_button.grid(row=3, columnspan=2, pady=10)

    def show_register(self):
        self.clear_window()
        self.login_button.grid_remove()
        self.register_button.grid_remove()

        self.label_fullname = tk.Label(self, text="Full Name")
        self.label_fullname.grid(row=0, column=0)
        self.entry_fullname = tk.Entry(self)
        self.entry_fullname.grid(row=0, column=1)
        self.label_password = tk.Label(self, text="Password")
        self.label_password.grid(row=1, column=0)
        self.entry_password = tk.Entry(self, show="*")
        self.entry_password.grid(row=1, column=1)
        self.label_email = tk.Label(self, text="Email")
        self.label_email.grid(row=2, column=0)
        self.entry_email = tk.Entry(self)
        self.entry_email.grid(row=2, column=1)
        self.register_button = tk.Button(self, text="Register", command=self.register_user)
        self.register_button.grid(row=3, columnspan=2, pady=10)
        back_button = tk.Button(self, text="Back", command=self.redirect_main_menu_or_login)
        back_button.grid(row=4, columnspan=2, pady=10)

    def hide_login_register(self):
        self.login_button.grid_remove()
        self.register_button.grid_remove()

    def clear_window(self):
        for widget in self.winfo_children():
            widget.grid_remove()

    def redirect_login_page(self):
        self.show_login()

    def login_user(self):
        username = self.entry_username.get()
        password = self.entry_password.get()

        if not full_name_validator(username):
            print("Invalid full name. Please provide a valid name.")
            return

        if not password_validator(password):
            print("Invalid password. Password should be at least 8 characters long.")
            return

        users_data = pd.read_excel('users.xlsx')

        user_match = users_data[((users_data['full_name'] == username) | (users_data['email'] == username)) &
                                (users_data['password'] == password)]

        if not user_match.empty:
            print("Logged in!")
            user_index = user_match.index[0]
            user_ban_status = user_match.at[user_index, 'is_ban']
            if user_ban_status == 1:
                print("Sorry, your account is banned.")
                return
            user_session_data = users_data.loc[users_data['full_name'] == username]
            self.user_session = {
                "name": username,
                "permissions": user_session_data.iloc[0]['permissions']
            }
            self.redirect_main_menu()
        else:
            print("Invalid credentials. Please register or try again.")
            self.redirect_login_page()

    def register_user(self):
        fullname = self.entry_fullname.get()
        password = self.entry_password.get()
        email = self.entry_email.get()

        check_users_data = pd.read_excel('users.xlsx')
        if check_users_data[(check_users_data['full_name'] == fullname) | (check_users_data['email'] == email)].shape[0] > 0:
            print("User with the same full name or email already exists.")
            return

        if not email_validator(email):
            print("Invalid email format. Please enter a valid email.")
            return

        if not full_name_validator(fullname):
            print("Invalid full name. Please provide a valid name.")
            return

        if not password_validator(password):
            print("Invalid password. Password should be at least 8 characters long.")
            return
        users_data = pd.read_excel('users.xlsx')
        new_user = pd.DataFrame({'id': uuid.uuid4(), 'full_name': [fullname], 'password': [password], 'email': [email],
                                 'create_at': [datetime.datetime.now()],
                                 'permissions': [0 if users_data.empty else 2],
                                 'is_ban': [0], 'borrowed_books': [dict()]})
        users_data = users_data._append(new_user, ignore_index=True)

        users_data.to_excel('users.xlsx', index=False)
        print("User registered successfully. Please login.")
        self.redirect_login_page()

    def redirect_main_menu(self):
        self.clear_window()
        self.hide_login_register()
        self.main_menu_label = tk.Label(self, text="Welcome to the Main Menu")
        self.main_menu_label.grid(row=0, columnspan=2)
        self.show_books = tk.Button(self, text="Show Books", command=self.show_books)
        self.show_books.grid(row=1, column=0, pady=10)
        self.loan_book = tk.Button(self, text="Loan a Book", command=self.loan_a_book)
        self.loan_book.grid(row=1, column=1, pady=10)
        self.borrow_book = tk.Button(self, text="Borrow Book", command=self.borrow_book_func)
        self.borrow_book.grid(row=2, columnspan=2, pady=10)
        self.show_borrowed_books = tk.Button(self, text="Show Borrowed Books", command=self.show_borrowed_books)
        self.show_borrowed_books.grid(row=3, column=0, pady=10)
        self.show_loaned_books = tk.Button(self, text="Show Loaned Books", command=self.show_loaned_books)
        self.show_loaned_books.grid(row=3, column=1, pady=10)

        if self.user_session["permissions"] == 0:
            self.show_users_button = tk.Button(self, text="Show Users", command=self.show_users)
            self.show_users_button.grid(row=4, columnspan=2, pady=10)

            self.register_user_button = tk.Button(self, text="Register User", command=self.register_user_by_admin)
            self.register_user_button.grid(row=5, columnspan=2, pady=10)

            self.change_user_info_button = tk.Button(self, text="Change User Info", command=self.change_user_info)
            self.change_user_info_button.grid(row=6, columnspan=2, pady=10)

            self.delete_user_button = tk.Button(self, text="Delete User", command=self.delete_user)
            self.delete_user_button.grid(row=7, columnspan=2, pady=10)

            self.allow_disallow_button = tk.Button(self, text="Allow/Disallow Borrowing",
                                                   command=self.allow_disallow_borrow)
            self.allow_disallow_button.grid(row=8, columnspan=2, pady=10)

        if self.user_session["permissions"] in [0, 1]:
            self.ban_users_button = tk.Button(self, text="Ban Users", command=self.ban_users)
            self.ban_users_button.grid(row=9, columnspan=2, pady=10)

            self.register_user_staff_button = tk.Button(self, text="Register User (Normal)",
                                                        command=self.register_user_staff)
            self.register_user_staff_button.grid(row=10, columnspan=2, pady=10)

    def delete_user(self):
        delete_user_window = tk.Toplevel(self)
        delete_user_window.title("Delete User")

        tk.Label(delete_user_window, text="Full Name").grid(row=0, column=0)
        entry_fullname = tk.Entry(delete_user_window)
        entry_fullname.grid(row=0, column=1)

        tk.Label(delete_user_window, text="Email").grid(row=1, column=0)
        entry_email = tk.Entry(delete_user_window)
        entry_email.grid(row=1, column=1)

        def remove_user():
            fullname = entry_fullname.get()
            email = entry_email.get()
            if not email_validator(email):
                print("Invalid email format. Please enter a valid email.")
                return

            if not full_name_validator(fullname):
                print("Invalid full name. Please provide a valid name.")
                return

            users_data = pd.read_excel('users.xlsx')
            user_match = users_data[
                (users_data['full_name'] == fullname) & (users_data['email'] == email)
                ]

            if not user_match.empty:
                user_index = user_match.index[0]
                borrowed_books = users_data.loc[user_index, 'borrowed_books']

                if pd.isnull(borrowed_books) or borrowed_books == '{}':
                    users_data.drop(user_index, inplace=True)
                    users_data.to_excel('users.xlsx', index=False)
                    print("User deleted successfully.")
                    delete_user_window.destroy()
                else:
                    print(
                        "This user has borrowed books from our library. Please ask the user to return the books or consider banning the user.")
            else:
                print("User not found. Please check the entered details.")

        delete_button = tk.Button(delete_user_window, text="Delete User", command=remove_user)
        delete_button.grid(row=2, columnspan=2)

    def allow_disallow_borrow(self):
        allow_disallow_window = tk.Toplevel(self)
        allow_disallow_window.title("Allow/Disallow Borrowing")

        tk.Label(allow_disallow_window, text="Book Name").grid(row=0, column=0)
        entry_book_name = tk.Entry(allow_disallow_window)
        entry_book_name.grid(row=0, column=1)

        tk.Label(allow_disallow_window, text="Author").grid(row=1, column=0)
        entry_book_author = tk.Entry(allow_disallow_window)
        entry_book_author.grid(row=1, column=1)

        tk.Label(allow_disallow_window, text="Allow Borrowed?").grid(row=2, column=0)
        allow_borrow_options = ["YES", "NO"]
        allow_borrow_var = tk.StringVar(allow_disallow_window)
        allow_borrow_var.set(allow_borrow_options[0])
        allow_borrow_dropdown = tk.OptionMenu(allow_disallow_window, allow_borrow_var, *allow_borrow_options)
        allow_borrow_dropdown.grid(row=2, column=1)

        def modify_allow_borrow():
            book_name = entry_book_name.get()
            book_author = entry_book_author.get()
            allow_borrow_choice = allow_borrow_var.get()

            books_data = pd.read_excel('books.xlsx')
            book_match = books_data[
                (books_data['name'] == book_name) & (books_data['author'] == book_author)
                ]

            if not book_match.empty:
                book_index = book_match.index[0]
                books_data.at[book_index, 'allow_borrowed'] = allow_borrow_choice

                books_data.to_excel('books.xlsx', index=False)
                print(f"Allow Borrowed status updated for {book_name} successfully.")
                allow_disallow_window.destroy()
            else:
                print("Book not found. Please check the entered details.")

        submit_button = tk.Button(allow_disallow_window, text="Modify Allow Borrowed", command=modify_allow_borrow)
        submit_button.grid(row=3, columnspan=2)

    def ban_users(self):
        ban_users_window = tk.Toplevel(self)
        ban_users_window.title("Ban Users")

        tk.Label(ban_users_window, text="Full Name").grid(row=0, column=0)
        entry_fullname = tk.Entry(ban_users_window)
        entry_fullname.grid(row=0, column=1)

        tk.Label(ban_users_window, text="Email").grid(row=1, column=0)
        entry_email = tk.Entry(ban_users_window)
        entry_email.grid(row=1, column=1)

        tk.Label(ban_users_window, text="Is Banned?").grid(row=2, column=0)
        is_banned_options = ["Yes", "No"]
        is_banned_var = tk.StringVar(ban_users_window)
        is_banned_var.set(is_banned_options[0])
        is_banned_dropdown = tk.OptionMenu(ban_users_window, is_banned_var, *is_banned_options)
        is_banned_dropdown.grid(row=2, column=1)

        def modify_banned_status():
            fullname = entry_fullname.get()
            email = entry_email.get()
            is_banned_choice = is_banned_var.get()
            if not email_validator(email):
                print("Invalid email format. Please enter a valid email.")
                return

            if not full_name_validator(fullname):
                print("Invalid full name. Please provide a valid name.")
                return

            users_data = pd.read_excel('users.xlsx')
            user_match = users_data[
                (users_data['full_name'] == fullname) & (users_data['email'] == email)
                ]

            if not user_match.empty:
                user_index = user_match.index[0]
                is_banned_code = 1 if is_banned_choice == "Yes" else 0
                users_data.at[user_index, 'is_ban'] = is_banned_code

                users_data.to_excel('users.xlsx', index=False)
                print(f"Banned status updated for {fullname} successfully.")
                ban_users_window.destroy()
            else:
                print("User not found. Please check the entered details.")

        submit_button = tk.Button(ban_users_window, text="Modify Banned Status", command=modify_banned_status)
        submit_button.grid(row=3, columnspan=2)

    def register_user_staff(self):
        register_staff_window = tk.Toplevel(self)
        register_staff_window.title("Register User by Staff")

        tk.Label(register_staff_window, text="Full Name").grid(row=0, column=0)
        entry_fullname = tk.Entry(register_staff_window)
        entry_fullname.grid(row=0, column=1)

        tk.Label(register_staff_window, text="Password").grid(row=1, column=0)
        entry_password = tk.Entry(register_staff_window, show="*")
        entry_password.grid(row=1, column=1)

        tk.Label(register_staff_window, text="Email").grid(row=2, column=0)
        entry_email = tk.Entry(register_staff_window)
        entry_email.grid(row=2, column=1)

        def add_user_staff():
            fullname = entry_fullname.get()
            password = entry_password.get()
            email = entry_email.get()

            check_users_data = pd.read_excel('users.xlsx')
            if check_users_data[(check_users_data['full_name'] == fullname) | (check_users_data['email'] == email)].shape[0] > 0:
                print("User with the same full name or email already exists.")
                return

            if not email_validator(email):
                print("Invalid email format. Please enter a valid email.")
                return

            if not full_name_validator(fullname):
                print("Invalid full name. Please provide a valid name.")
                return

            if not password_validator(password):
                print("Invalid password. Password should be at least 8 characters long.")
                return
            users_data = pd.read_excel('users.xlsx')
            new_user = pd.DataFrame(
                {'id': uuid.uuid4(), 'full_name': [fullname], 'password': [password], 'email': [email],
                 'create_at': [datetime.datetime.now()],
                 'permissions': [2],
                 'is_ban': [0], 'borrowed_books': [dict()]})
            users_data = users_data._append(new_user, ignore_index=True)

            users_data.to_excel('users.xlsx', index=False)
            print("User registered successfully by staff.")
            register_staff_window.destroy()

        submit_button = tk.Button(register_staff_window, text="Register User", command=add_user_staff)
        submit_button.grid(row=3, columnspan=2)

    def change_user_info(self):
        change_user_window = tk.Toplevel(self)
        change_user_window.title("Change User Information")

        tk.Label(change_user_window, text="ID").grid(row=0, column=0)
        entry_id = tk.Entry(change_user_window)
        entry_id.grid(row=0, column=1)

        tk.Label(change_user_window, text="Full Name").grid(row=1, column=0)
        entry_fullname = tk.Entry(change_user_window)
        entry_fullname.grid(row=1, column=1)

        tk.Label(change_user_window, text="New Email").grid(row=2, column=0)
        entry_new_email = tk.Entry(change_user_window)
        entry_new_email.grid(row=2, column=1)

        tk.Label(change_user_window, text="New Password").grid(row=3, column=0)
        entry_new_password = tk.Entry(change_user_window, show="*")
        entry_new_password.grid(row=3, column=1)

        tk.Label(change_user_window, text="New Permissions").grid(row=4, column=0)
        permissions_options = ["Super Admin", "Staff", "Normal Client"]
        permissions_var = tk.StringVar(change_user_window)
        permissions_var.set(permissions_options[0])  # default value
        permissions_dropdown = tk.OptionMenu(change_user_window, permissions_var, *permissions_options)
        permissions_dropdown.grid(row=4, column=1)

        def modify_user_info():
            id = entry_id.get()
            fullname = entry_fullname.get()
            new_email = entry_new_email.get()
            new_password = entry_new_password.get()
            new_permissions = permissions_var.get()
            if not email_validator(new_email):
                print("Invalid email format. Please enter a valid email.")
                return

            if not full_name_validator(fullname):
                print("Invalid full name. Please provide a valid name.")
                return

            if not password_validator(new_password):
                print("Invalid password. Password should be at least 8 characters long.")
                return

            permissions_map = {"Super Admin": 0, "Staff": 1, "Normal Client": 2}
            new_permissions_code = permissions_map.get(new_permissions)

            users_data = pd.read_excel('users.xlsx')
            user_match = users_data[users_data['id'] == id]

            if not user_match.empty:
                user_index = user_match.index[0]
                users_data.at[user_index, 'fullname'] = fullname
                users_data.at[user_index, 'email'] = new_email
                users_data.at[user_index, 'password'] = new_password
                users_data.at[user_index, 'permissions'] = new_permissions_code

                users_data.to_excel('users.xlsx', index=False)
                print("User information updated successfully.")
                change_user_window.destroy()
            else:
                print("User not found. Please check the entered details.")

        submit_button = tk.Button(change_user_window, text="Modify User Info", command=modify_user_info)
        submit_button.grid(row=5, columnspan=2)

    def register_user_by_admin(self):
        register_user_window = tk.Toplevel(self)
        register_user_window.title("Register User by Admin")

        tk.Label(register_user_window, text="Full Name").grid(row=0, column=0)
        entry_fullname = tk.Entry(register_user_window)
        entry_fullname.grid(row=0, column=1)

        tk.Label(register_user_window, text="Password").grid(row=1, column=0)
        entry_password = tk.Entry(register_user_window, show="*")
        entry_password.grid(row=1, column=1)

        tk.Label(register_user_window, text="Email").grid(row=2, column=0)
        entry_email = tk.Entry(register_user_window)
        entry_email.grid(row=2, column=1)

        tk.Label(register_user_window, text="Permissions").grid(row=3, column=0)
        permissions_options = ["Super Admin", "Staff", "Normal Client"]
        permissions_var = tk.StringVar(register_user_window)
        permissions_var.set(permissions_options[0])
        permissions_dropdown = tk.OptionMenu(register_user_window, permissions_var, *permissions_options)
        permissions_dropdown.grid(row=3, column=1)

        def submit_registration():
            fullname = entry_fullname.get()
            password = entry_password.get()
            email = entry_email.get()
            permissions = permissions_var.get()

            check_users_data = pd.read_excel('users.xlsx')
            if check_users_data[(check_users_data['full_name'] == fullname) | (check_users_data['email'] == email)].shape[0] > 0:
                print("User with the same full name or email already exists.")
                return

            if not email_validator(email):
                print("Invalid email format. Please enter a valid email.")
                return

            if not full_name_validator(fullname):
                print("Invalid full name. Please provide a valid name.")
                return

            if not password_validator(password):
                print("Invalid password. Password should be at least 8 characters long.")
                return
            permissions_map = {"Super Admin": 0, "Staff": 1, "Normal Client": 2}
            permissions_code = permissions_map.get(permissions)

            users_data = pd.read_excel('users.xlsx')
            new_user = pd.DataFrame({
                'id': [str(uuid.uuid4())],
                'full_name': [fullname],
                'password': [password],
                'email': [email],
                'create_at': [datetime.datetime.now()],
                'permissions': [permissions_code],
                'is_ban': [0],
                'borrowed_books': [{}]
            })

            users_data = users_data._append(new_user, ignore_index=True)
            users_data.to_excel('users.xlsx', index=False)
            print("User registered successfully.")
            register_user_window.destroy()

        submit_button = tk.Button(register_user_window, text="Register", command=submit_registration)
        submit_button.grid(row=4, columnspan=2)

    def show_users(self):
        users_data = pd.read_excel('users.xlsx')

        user_table = tk.Toplevel(self)
        user_table.title("Users Information")

        columns = ['full_name', 'email', 'create_at', 'borrowed_books', 'is_ban', 'permissions']

        for col_index, col_name in enumerate(columns):
            header_label = tk.Label(user_table, text=col_name.capitalize())
            header_label.grid(row=0, column=col_index)

        for index, row in users_data.iterrows():
            if row['is_ban'] == 0:
                is_ban_value = 'NO'
            else:
                is_ban_value = 'YES'

            if row['permissions'] == 0:
                del row['permissions']
                row['permissions'] = 'super admin'
            elif row['permissions'] == 1:
                del row['permissions']
                row['permissions'] = 'staff'
            elif row['permissions'] == 2:
                del row['permissions']
                row['permissions'] = 'normal client'

            data_row = []
            for col in columns:
                if col != 'is_ban':
                    data_row.append(row[col])
                elif col == 'is_ban':
                    data_row.append(is_ban_value)

            for col_index, col_value in enumerate(data_row):
                label = tk.Label(user_table, text=col_value)
                label.grid(row=index + 1, column=col_index)

        def close_user_table():
            user_table.destroy()

        close_button = tk.Button(user_table, text="Close", command=close_user_table)
        close_button.grid(row=index + 2, columnspan=len(columns))

    def show_borrowed_books(self):
        users_data = pd.read_excel('users.xlsx')
        borrowed_books = users_data.loc[users_data['full_name'] == self.user_session["name"], 'borrowed_books'].iloc[0]
        borrowed_books = ast.literal_eval(borrowed_books)
        if len(borrowed_books) == 0:
            print("You haven't borrowed any books yet.")
            return
        borrowed_books_table = tk.Toplevel(self)
        borrowed_books_table.title("Borrowed Books")

        row = 0
        for book, count in borrowed_books.items():
            tk.Label(borrowed_books_table, text="Book Name: ").grid(row=row, column=0)
            tk.Label(borrowed_books_table, text=book[0]).grid(row=row, column=1)
            tk.Label(borrowed_books_table, text="Author: ").grid(row=row + 1, column=0)
            tk.Label(borrowed_books_table, text=book[1]).grid(row=row + 1, column=1)
            tk.Label(borrowed_books_table, text="Count: ").grid(row=row + 2, column=0)
            tk.Label(borrowed_books_table, text=count).grid(row=row + 2, column=1)
            row += 3

        def close_borrowed_books():
            borrowed_books_table.destroy()

        close_button = tk.Button(borrowed_books_table, text="Close", command=close_borrowed_books)
        close_button.grid(row=row, columnspan=2)


    def show_loaned_books(self):
        books_data = pd.read_excel('books.xlsx')
        user_loaned_books = books_data.loc[books_data['user_loaned'].apply(
            lambda x: self.user_session["name"] in ast.literal_eval(x) if pd.notnull(x) else False)]

        if not user_loaned_books.empty:
            user_loaned_books_table = tk.Toplevel(self)
            user_loaned_books_table.title("Loaned Books")


            row = 0
            for _, book in user_loaned_books.iterrows():
                tk.Label(user_loaned_books_table, text="Book Name: ").grid(row=row, column=0)
                tk.Label(user_loaned_books_table, text=book['name']).grid(row=row, column=1)
                tk.Label(user_loaned_books_table, text="Count: ").grid(row=row + 1, column=0)
                tk.Label(user_loaned_books_table,
                         text=ast.literal_eval(book['user_loaned'])[self.user_session["name"]]).grid(
                    row=row + 1,
                    column=1)
                row += 2

            def close_loaned_books():
                user_loaned_books_table.destroy()

            close_button = tk.Button(user_loaned_books_table, text="Close", command=close_loaned_books)
            close_button.grid(row=row, columnspan=2)
        else:
            print("You haven't loaned any books yet.")

    def borrow_book_func(self):
        modal_window = tk.Toplevel(self)
        modal_window.title("Borrow a Book")

        label_borrow_book = tk.Label(modal_window, text="Borrow a Book")
        label_borrow_book.grid(row=0, columnspan=2)

        label_book_name = tk.Label(modal_window, text="Name of the Book")
        label_book_name.grid(row=1, column=0)
        entry_book_name = tk.Entry(modal_window)
        entry_book_name.grid(row=1, column=1)

        label_book_author = tk.Label(modal_window, text="Author of Book")
        label_book_author.grid(row=2, column=0)
        entry_book_author = tk.Entry(modal_window)
        entry_book_author.grid(row=2, column=1)

        label_book_count = tk.Label(modal_window, text="Number of Copies")
        label_book_count.grid(row=3, column=0)
        entry_book_count = tk.Entry(modal_window)
        entry_book_count.grid(row=3, column=1)

        borrow_button = tk.Button(modal_window, text="Borrow Book", command=lambda: self.borrow_book_action(
            entry_book_name.get(), entry_book_author.get(), entry_book_count.get(), modal_window))
        borrow_button.grid(row=4, columnspan=2, pady=10)

    def borrow_book_action(self, book_name, book_author, book_count, modal_window):
        try:
            book_count = int(book_count)
        except ValueError:
            print("Not a valid book!")
            modal_window.destroy()
            return

        books_data = pd.read_excel('books.xlsx')
        users_data = pd.read_excel('users.xlsx')

        book_index = books_data[(books_data['name'] == book_name) & (books_data['author'] == book_author)].index

        if book_index.empty:
            print("This book does not exist.")
            return

        book_row_index = book_index[0]

        book_row = books_data.loc[book_row_index]

        if book_row['allow_borrowed'] == 'NO':
            print("This book is not available for borrowing.")
            return

        if book_row['count'] < book_count:
            print("Not enough copies available to borrow.")
            return

        user_index = users_data[users_data['full_name'] == self.user_session["name"]].index[0]

        current_borrowed_books = users_data.at[user_index, 'borrowed_books']

        if current_borrowed_books:
            if isinstance(current_borrowed_books, float) and pd.isna(current_borrowed_books):
                current_borrowed_books = {}
            else:
                current_borrowed_books = eval(current_borrowed_books)
            if (book_name, book_author) in current_borrowed_books:
                current_borrowed_books[(book_name, book_author)] += book_count
            else:
                current_borrowed_books[(book_name, book_author)] = book_count
        else:
            current_borrowed_books = {(book_name, book_author): book_count}

        users_data.at[user_index, 'borrowed_books'] = str(current_borrowed_books)
        books_data.at[book_row_index, 'count'] -= book_count

        if books_data.at[book_row_index, 'user_loaned'] and self.user_session["name"] in books_data.at[
            book_row_index, 'user_loaned']:
            user_loaned_books = eval(books_data.at[book_row_index, 'user_loaned'])
            if self.user_session["name"] in user_loaned_books:
                user_loaned_books[self.user_session["name"]] -= book_count
                if user_loaned_books[self.user_session["name"]] == 0:
                    del user_loaned_books[self.user_session["name"]]
                elif user_loaned_books[self.user_session["name"]] < 0:

                    current_borrowed = users_data.at[user_index, 'borrowed_books']
                    if current_borrowed:
                        current_borrowed = eval(current_borrowed)
                        current_borrowed[(book_name, book_author)] = abs(user_loaned_books[self.user_session["name"]])
                        users_data.at[user_index, 'borrowed_books'] = str(current_borrowed)
                    else:
                        users_data.at[user_index, 'borrowed_books'] = str(
                            {book_name: abs(user_loaned_books[self.user_session["name"]])})
                    del user_loaned_books[self.user_session["name"]]

                books_data.at[book_row_index, 'user_loaned'] = str(user_loaned_books)

        users_data.to_excel('users.xlsx', index=False)
        books_data.to_excel('books.xlsx', index=False)
        modal_window.destroy()
        print("Book borrowed successfully!")

    def show_books(self):
        books_data = pd.read_excel('books.xlsx')

        book_table_window = tk.Toplevel(self)
        book_table_window.title("Book Information")

        columns = ['name', 'author', 'count', 'allow_borrowed', 'user_loaned']
        book_info = books_data[columns]

        book_table_label = tk.Label(book_table_window, text="Book Information")
        book_table_label.grid(row=0, columnspan=len(columns))

        search_frame = tk.Frame(book_table_window)
        search_frame.grid(row=1, columnspan=len(columns))

        search_label = tk.Label(search_frame, text="Search by Name or Author:")
        search_label.grid(row=0, column=0)

        search_entry = tk.Entry(search_frame)
        search_entry.grid(row=0, column=1)

        def search_books():
            search_text = search_entry.get().lower()
            filtered_books = book_info[
                (book_info['name'].str.lower().str.contains(search_text)) |
                (book_info['author'].str.lower().str.contains(search_text))
                ]
            self.display_books(filtered_books, book_table_window, columns)

        search_button = tk.Button(search_frame, text="Search or Empty to show all books", command=search_books)
        search_button.grid(row=0, column=2)

    def display_books(self, book_info, window, columns):
        for widget in window.winfo_children():
            widget.destroy()

        book_table_label = tk.Label(window, text="Book Information")
        book_table_label.grid(row=0, columnspan=len(columns))

        for col_index, col_name in enumerate(columns):
            header_label = tk.Label(window, text=col_name.capitalize())
            header_label.grid(row=1, column=col_index)

        for index, row in book_info.iterrows():
            for col_index, col_name in enumerate(columns):
                label = tk.Label(window, text=row[col_name])
                label.grid(row=index + 2, column=col_index)

    def loan_a_book(self):
        loan_book_modal = tk.Toplevel(self)
        loan_book_modal.title("Loan a Book")

        label_name = tk.Label(loan_book_modal, text="Name of the Book")
        label_name.grid(row=0, column=0)
        entry_name = tk.Entry(loan_book_modal)
        entry_name.grid(row=0, column=1)

        label_author = tk.Label(loan_book_modal, text="Author of the Book")
        label_author.grid(row=1, column=0)
        entry_author = tk.Entry(loan_book_modal)
        entry_author.grid(row=1, column=1)

        label_count = tk.Label(loan_book_modal, text="Number of Copies")
        label_count.grid(row=2, column=0)
        entry_count = tk.Entry(loan_book_modal)
        entry_count.grid(row=2, column=1)

        borrow_button = tk.Button(
            loan_book_modal,
            text="Loan My Book or register a book",
            command=lambda: self.loan_my_book(
                entry_name.get(),
                entry_author.get(),
                entry_count.get(),
                loan_book_modal
            )
        )
        borrow_button.grid(row=4, columnspan=2, pady=10)

    def loan_my_book(self, name, author, count, modal):
        try:
            count = int(count)
        except ValueError:
            print("not a valid book")
            return
        allow_borrowed = "YES"
        books_data = pd.read_excel('books.xlsx')
        users_data = pd.read_excel('users.xlsx')
        book_index = books_data[(books_data['name'] == name) & (books_data['author'] == author)].index.tolist()
        user_index = users_data[users_data['full_name'] == self.user_session["name"]].index.tolist()
        if book_index:
            book_index = book_index[0]
            user_index = user_index[0]
            user_loaned = books_data.loc[book_index, 'user_loaned']
            borrowed_books = users_data.loc[user_index, 'borrowed_books']
            if isinstance(borrowed_books, float) and pd.isna(user_loaned):
                borrowed_books = {}
            else:
                borrowed_books = ast.literal_eval(borrowed_books)
            if (name, author) in borrowed_books:
                borrowed_count = borrowed_books[(name, author)]
                new_count = count - borrowed_count
                if new_count == 0:
                    del borrowed_books[(name, author)]
                    users_data.at[user_index, 'borrowed_books'] = str(borrowed_books)
                    users_data.to_excel('users.xlsx', index=False)
                    print("thanks for returning the book!")
                    return
                else:
                    print('Thanks for returning the book and extra copies!')
                    count = abs(new_count)
                    del borrowed_books[(name, author)]
                users_data.at[user_index, 'borrowed_books'] = str(borrowed_books)
            if isinstance(user_loaned, float) and pd.isna(user_loaned):
                user_loaned = {}  # Handling NaN values
            else:
                user_loaned = ast.literal_eval(user_loaned)
            if self.user_session["name"] in user_loaned:
                user_loaned[self.user_session["name"]] += count
            else:
                user_loaned[self.user_session["name"]] = count

            books_data.at[book_index, 'user_loaned'] = str(user_loaned)
            books_data.at[book_index, 'count'] += count
        else:
            new_book = pd.DataFrame({
                'id': [uuid.uuid4()],
                'name': [name],
                'author': [author],
                'count': [count],
                'allow_borrowed': [allow_borrowed],
                'user_loaned': [{self.user_session["name"]: count}]
            })
            books_data = books_data._append(new_book, ignore_index=True)
        users_data.to_excel('users.xlsx', index=False)
        books_data.to_excel('books.xlsx', index=False)
        modal.destroy()
        print("Book loaned successfully!")


if __name__ == "__main__":
    app = LibraryApp()
    app.redirect_main_menu_or_login()
    app.mainloop()

