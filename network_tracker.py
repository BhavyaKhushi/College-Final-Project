# network_tracker.py

from tkinter import *
import tkinter as tk
import tkinter.messagebox as mbox
from PIL import ImageTk, Image
import time
import psutil
import socket
import threading

def run_network_usage_tracker():
    def update_label():
        nonlocal old_value
        new_value = psutil.net_io_counters().bytes_sent + psutil.net_io_counters().bytes_recv
        usage = new_value - old_value
        usage_str = "{0:.3f}".format(usage)

        path_text.delete("1.0", "end")
        path_text.insert(END, "Usage : " + usage_str + " bytes/sec")

        IPaddress = socket.gethostbyname(socket.gethostname())
        if IPaddress == "127.0.0.1":
            l2.configure(text="No internet, your localhost is\n" + IPaddress)
        else:
            l2.configure(text="Connected, with the IP address\n" + IPaddress)

        if usage > 1000000:
            mbox.showinfo("Exceed Status", "Max Limit Usage Exceeded.")

        old_value = new_value
        window.after(1000, update_label)

    window = Tk()
    window.title("Network Usage Tracker")
    window.geometry("1000x700")

    Label(window, text="NETWORK USAGE\nTRACKER", font=("Arial", 50, 'underline'), fg="magenta").place(x=190, y=10)
    Label(window, text="MAX LIMIT  :  1 MB/sec", font=("Arial", 50), fg="green").place(x=130, y=180)

    global path_text
    path_text = Text(window, height=1, width=24, font=("Arial", 50), bg="white", fg="blue", borderwidth=2, relief="solid")
    path_text.place(x=50, y=300)

    Label(window, text="Connection Status :", font=("Arial", 50), fg="green").place(x=200, y=450)

    global l2
    l2 = Label(window, fg='blue', font=("Arial", 30))
    l2.place(x=200, y=530)

    def exit_win():
        if mbox.askokcancel("Exit", "Do you want to exit?"):
            window.destroy()

    window.protocol("WM_DELETE_WINDOW", exit_win)
    old_value = psutil.net_io_counters().bytes_sent + psutil.net_io_counters().bytes_recv

    update_label()
    window.mainloop()
