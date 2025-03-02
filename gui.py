from tkinter import Tk, Canvas, Entry, PhotoImage, filedialog, messagebox, StringVar, Toplevel, Button, Label, Frame
from tkinter.ttk import Combobox
from pathlib import Path
from datetime import datetime
import main
import webbrowser, os, subprocess, platform
import tkinter as tk

def show_api_key_window():
    global api_key_window
    
    if 'api_key_window' in globals() and api_key_window.winfo_exists():
        api_key_window.lift()  # Bring the existing window to the front
        return

    api_key_window = Toplevel(window)
    api_key_window.title("API Key")
    api_key_window.geometry("400x165")
    api_key_window.configure(bg="#CCCCCC")
    api_key_window.resizable(False, False) 

    # Center the new window on the main window
    main_window_width = window.winfo_width()
    main_window_height = window.winfo_height()
    main_window_x = window.winfo_x()
    main_window_y = window.winfo_y()

    new_window_width = 400
    new_window_height = 165

    new_window_x = main_window_x + (main_window_width // 2) - (new_window_width // 2)
    new_window_y = main_window_y + (main_window_height // 2) - (new_window_height // 2)

    api_key_window.geometry(f"{new_window_width}x{new_window_height}+{new_window_x}+{new_window_y}")

    api_key_label = Label(api_key_window, text="API Key:", bg="#CCCCCC", fg="black")
    api_key_label.place(x=50, y=20)

    api_key_entry = Entry(api_key_window, bd=0, bg="#EFF1F5", fg="#000716", highlightthickness=1, insertbackground="black", highlightcolor="#000716")
    api_key_entry.place(x=50, y=50, width=300, height=30)
    api_key_entry.insert(0, main.API_KEY)

    def update_api_key():
        new_api_key = api_key_entry.get()
        if new_api_key:
            main.update_api_key(new_api_key)
            api_key_window.destroy()
        else:
            messagebox.showerror("Error", "API key cannot be empty!")

    update_button = Button(api_key_window, text="Save", command=update_api_key, bd=0, cursor="hand2")
    update_button.place(x=150, y=100, width=100, height=30)

    def open_2captcha_link(event):
        webbrowser.open("https://2captcha.com/")

    link_frame = Frame(api_key_window, bg="#CCCCCC")
    link_frame.place(relx=0.5, y=150, anchor="center")

    link_label_prefix = Label(link_frame, text="visit ", fg="black", bg="#CCCCCC")
    link_label_prefix.pack(side="left")

    link_label_2captcha = Label(link_frame, text="2Captcha", fg="blue", bg="#CCCCCC", cursor="hand2")
    link_label_2captcha.pack(side="left")
    link_label_2captcha.bind("<Button-1>", open_2captcha_link)

    link_label_suffix = Label(link_frame, text=" for API key", fg="black", bg="#CCCCCC")
    link_label_suffix.pack(side="left")

def excel():
    
    entry_2.delete(0, "end")
    entry_2.insert(0, filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")]))
    window.focus_force()

def open_excel_file():
    global file_path
    file_path = entry_2.get()
    if file_path and Path(file_path).exists():
        if platform.system() == "Windows":
            os.startfile(file_path)
        elif platform.system() == "Darwin":  # macOS
            subprocess.call(("open", file_path))
        else:
            messagebox.showerror("Error", "Unsupported platform!")
    else:
        messagebox.showerror("Error", "No valid Excel file selected!")
        window.focus_force()

def dwnld():
    global folder_path
    entry_3.delete(0, "end")
    entry_3.insert(0, filedialog.askdirectory())
    folder_path = entry_3.get()
    window.focus_force()

def open_folder():
    folder = entry_3.get()
    if folder and Path(folder).exists():
        if platform.system() == "Windows":
            os.startfile(folder)
        elif platform.system() == "Darwin":
            subprocess.call(("open", folder))
        else:
            messagebox.showerror("Error", "Unsupported platform!")
    else:
        messagebox.showerror("Error", "No valid folder path selected!")
        window.focus_force()


def main_func(period, entry_2, entry_3, captcha_var):
    if period == "Quarterly":
        if entry_2.get() and entry_3.get():
            quarter()
            if year():
                window.iconify()
                main.main(entry_2.get(), entry_3.get(), period, captcha_var, window)
        else:
            messagebox.showinfo("Error", "Make sure to Select Excel File and Download Path before clicking Run.")
            window.focus_force()
    elif period == "Monthly":
        if entry_2.get() and entry_3.get():
            monthly()
            if year():
                window.iconify()
                main.main(entry_2.get(), entry_3.get(), period, captcha_var,window)
        else:
            messagebox.showinfo("Error", "Make sure to Select Excel File and Download Path before clicking Run.")
            window.focus_force()
    elif period == "All":
        if entry_2.get() and entry_3.get():
            all()
            if year():
                window.iconify()
                main.main(entry_2.get(), entry_3.get(), period, captcha_var, window)
        else:
            messagebox.showinfo("Error", "Make sure to Select Excel File and Download Path before clicking Run.")
            window.focus_force()
    else:
        messagebox.showinfo("Error", "Select a valid period to download.")
        window.focus_force()

def update_combo(*args):
    period = combo_var1.get()
    combo_2.place_forget()
    combo_3.place_forget()
    if period == "Monthly":
        combo_values = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        combo_2.config(values=combo_values, state="readonly")
        combo_2.place(x=75.5, y=430.0, width=190.0, height=25.0)
    elif period == "Quarterly":
        combo_values = ["Quarter 1 (Apr - Jun)", "Quarter 2 (Jul - Sep)", "Quarter 3 (Oct - Dec)", "Quarter 4 (Jan - Mar)"]
        combo_3.config(values=combo_values, state="readonly")
        combo_3.place(x=75.5, y=430.0, width=190.0, height=25.0)

def decide(period, entry_2, entry_3):
    captcha_val = captcha_var.get()
    main_func(period, entry_2, entry_3, captcha_val)

def year():
    selected_year = combo_var0.get()
    if selected_year != 'Year':
        main.year(selected_year)
        return True
    else:
        messagebox.showinfo("Error", "Please select an appropriate year")
        window.focus_force()
        return False

def quarter():
    selected_option = combo_var3.get()
    main.quarter(selected_option)

def monthly():
    selected_option = combo_var2.get()
    main.monthly(selected_option)

def all():
    main.all()

def round_rectangle(canvas, x1, y1, x2, y2, r=25, **kwargs):
    points = [
        x1+r, y1,
        x2-r, y1,
        x2, y1,
        x2, y1+r,
        x2, y2-r,
        x2, y2,
        x2-r, y2,
        x1+r, y2,
        x1, y2,
        x1, y2-r,
        x1, y1+r,
        x1, y1,
    ]
    return canvas.create_polygon(points, **kwargs, smooth=True)

def create_rounded_button(canvas, x1, y1, x2, y2, text, command):
    button_id = round_rectangle(canvas, x1, y1, x2, y2, r=20, fill="#000000", outline="")
    text_id = canvas.create_text((x1 + x2) // 2, (y1 + y2) // 2, text=text, font=("Bold", 15), fill="#ffffff")

    def on_click(event):
        command()

    def on_enter(event):
        canvas.itemconfig(button_id, fill="#FFFFFF")
        canvas.itemconfig(text_id, fill="#000000")
        window.config(cursor="hand2")

    def on_leave(event):
        canvas.itemconfig(button_id, fill="#000000")
        canvas.itemconfig(text_id, fill="#EFF1F5")
        window.config(cursor="")

    canvas.tag_bind(button_id, "<Button-1>", on_click)
    canvas.tag_bind(text_id, "<Button-1>", on_click)
    canvas.tag_bind(button_id, "<Enter>", on_enter)
    canvas.tag_bind(text_id, "<Enter>", on_enter)
    canvas.tag_bind(button_id, "<Leave>", on_leave)
    canvas.tag_bind(text_id, "<Leave>", on_leave)

    return button_id

def create_hoverable_text(canvas, x, y, text, command):
    text_id = canvas.create_text(x, y, text=text, font=("inter", 14), fill="DodgerBlue", anchor="nw")

    def on_click(event):
        command()

    def on_enter(event):
        canvas.itemconfig(text_id, font=("inter", 15, "underline"))
        window.config(cursor="hand2")


    def on_leave(event):
        canvas.itemconfig(text_id, font=("inter", 15))
        window.config(cursor="")

    canvas.tag_bind(text_id, "<Button-1>", on_click)
    canvas.tag_bind(text_id, "<Enter>", on_enter)
    canvas.tag_bind(text_id, "<Leave>", on_leave)

    return text_id


def create_entry_with_image(canvas, x, y, width, height, image_path):
    entry_frame = Canvas(canvas, bd=0, highlightthickness=0)
    entry_frame.place(x=x, y=y, width=width, height=height)

    entry_image = PhotoImage(file=image_path)
    entry_bg = entry_frame.create_image(width // 2, height // 2, image=entry_image)

    entry = Entry(entry_frame, bd=0, bg="#EFF1F5", fg="#000716", highlightthickness=0, state="readonly")
    entry.place(x=5, y=0, relwidth=0.94, height=height)

    return entry

window = Tk()
window.title("2b-Downloader")
window.geometry("800x600")
window.configure(bg="#efefef")

canvas = Canvas(window, height=600, width=800, relief="ridge", bg="#efefef")
canvas.place(x=0, y=0)

image_image_1 = PhotoImage(file="assets/frame0/image.png")
image_1 = canvas.create_image(760.0, 350.0, image=image_image_1)

current_year = datetime.now().year
combo_values = [f"{year}-{str(year+1)[-2:]}" for year in range(current_year, 2016, -1)]
combo_var0 = StringVar()
combo_var0.set("Year")
combo_var0.trace_add("write", update_combo)
combo_0 = Combobox(window, values=combo_values, textvariable=combo_var0, state="readonly")
combo_0.place(x=275.0, y=372.0, width=125.0, height=25.0)
combo_0.config(font=('Inter', 14))

combo_values = ["Monthly", "Quarterly", "All"]
combo_var1 = StringVar()
combo_var1.set("Select")
combo_var1.trace_add("write", update_combo)
combo_1 = Combobox(window, values=combo_values, textvariable=combo_var1, state="readonly")
combo_1.place(x=75.5, y=372.0, width=125.0, height=25.0)
combo_1.config(font=('Inter', 14))

selected_period = combo_var1.get()

combo_values_2 = ["Select"]
combo_var2 = StringVar()
combo_var2.set("Select")
combo_2 = Combobox(window, values=combo_values_2, textvariable=combo_var2, state="readonly")

combo_values_3 = ["Select"]
combo_var3 = StringVar()
combo_var3.set("Select")
combo_3 = Combobox(window, values=combo_values_3, textvariable=combo_var3, state="readonly")

entry_image_2 = PhotoImage(file=("assets/frame0/entry_2.png"))
entry_bg_2 = canvas.create_image(
    202.5,
    287.5,
    image=entry_image_2
)
entry_2 = Entry(
    bd=0,
    bg="#EFF1F5",
    fg="#000716",
    highlightthickness=0,
    insertbackground="black"
)
entry_2.pack(fill='both', expand=True)
entry_2.place(
    x=71.0,
    y=170.0,
    width=263.0,
    height=29.0
)

entry_image_3 = PhotoImage(file=("assets/frame0/entry_2.png"))
entry_bg_3 = canvas.create_image(
    202.5,
    187.5,
    image=entry_image_3
)
entry_3 = Entry(
    bd=0,
    bg="#EFF1F5",
    fg="#000716",
    highlightthickness=0,
    insertbackground="black"
)
entry_3.pack(fill='both', expand=True)
entry_3.place(
    x=71.0,
    y=270.0,
    width=263.0,
    height=29.0
)

canvas.create_text(60.0, 130.0, anchor="nw", text=".xlsx file", fill="#000000", font=("Inter SemiBold", 16 * -1))
canvas.create_text(60.0, 230.0, anchor="nw", text="Download path", fill="#000000", font=("Inter SemiBold", 16 * -1))
canvas.create_rectangle(0.0, 0.0, 800.0, 80.0, fill="#1C1B1F", outline="")
canvas.create_text(12.0, 32.0, anchor="nw", text="GST-2B Downloader", fill="#FFFFFF", font=("Inter SemiBold", 24 * -1))

create_hoverable_text(canvas, 725, 90, "API Key", show_api_key_window)
create_rounded_button(canvas, 408, 165, 519, 204, "Select File", excel)
create_hoverable_text(canvas, 540, 175, "Open .xlsx", open_excel_file)
create_rounded_button(canvas, 408, 265, 519, 304, "Select Path", dwnld)
create_hoverable_text(canvas, 540, 275, "Open Folder", open_folder)
create_rounded_button(canvas, 595, 510, 759, 560, "Run", lambda: decide(combo_var1.get(), entry_2, entry_3))

captcha_var = tk.StringVar(value="auto")

radio_button_1 = tk.Radiobutton(canvas, text="Auto Captcha", variable=captcha_var, value="auto", bg=window.cget('bg'), fg='black', highlightthickness=0, bd=0,font=("inter", 14))
radio_button_2 = tk.Radiobutton(canvas, text="Manual Captcha", variable=captcha_var, value="manual", bg=window.cget('bg'),fg='black', highlightthickness=0, bd=0, font=("inter", 14))
radio_button_3 = tk.Radiobutton(canvas, text="ML Captcha", variable=captcha_var, value="ML", bg=window.cget('bg'),fg='black', highlightthickness=0, bd=0, font=("inter", 14))

# Place radio buttons on the canvas
canvas.create_window(650, 30, window=radio_button_1, anchor="nw")
canvas.create_window(650, 60, window=radio_button_2, anchor="nw")

# Place radio buttons on the canvas
radio_button_1.place(x=75.5, y=485)
radio_button_2.place(x=275, y=485)
radio_button_3.place(x=475, y=485)

window.resizable(False, False)
window.mainloop()