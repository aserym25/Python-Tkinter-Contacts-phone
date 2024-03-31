import tkinter as tk
import win32com.client

def copy_sim_contacts():
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    contacts_folder = namespace.GetDefaultFolder(10)  # 10 represents the Contacts folder

    sim_contacts = []
    for contact in contacts_folder.Items:
        if contact.MobileTelephoneNumber:
            sim_contacts.append(contact.MobileTelephoneNumber)

    # Copy the contacts to the clipboard
    root.clipboard_clear()
    root.clipboard_append('\n'.join(sim_contacts))
    root.update()

    # Display a message box
    tk.messagebox.showinfo("Success", "SIM contacts copied to clipboard!")

# Create the Tkinter window
root = tk.Tk()
root.title("SIM Contacts")
root.geometry("300x100")

# Create a button to copy the contacts
copy_button = tk.Button(root, text="Copy SIM Contacts", command=copy_sim_contacts)
copy_button.pack(pady=20)

# Start the Tkinter event loop
root.mainloop()