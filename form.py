import tkinter as tk

def submit():
    name = entry_name.get()
    age = entry_age.get()
    print("Name", name)
    print("Age", age)

    root = tk.Tk()
    root.title("Data Entry Form")

    label_name = tk.Label(root, text = "Name")
    label_name.grid(row=0, column=0, padx=10, pady=5)
    entry_name = tk.Entry(root)
    entry_name.grid(row=0, column=1, padx=10, pady=5)

    label_age = tk.Label(root, text = "Age")
    label_age.grid(row=1, column=0, padx=10, pady=5)
    entry_age = tk.Entry(root)
    entry_age.grid(row=1, column=1, padx=10, pady=5)

    submit_button = tk.Button(root, text = "Submit", command=submit)
    submit_button.grid(row=2, column=2, padx=10, pady=10)

    root.mainloop()