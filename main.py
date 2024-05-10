import os
import sys
import threading
import time
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from tkinter.filedialog import asksaveasfile
import serial
import serial.tools.list_ports
import datetime
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg)
from matplotlib.figure import Figure


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def check_is_number(text):
    try:
        float(text)
        return True
    except:
        return False


ser = serial.Serial()

baudrate_list = [150, 300, 1200, 2400, 4800, 9600, 19200, 38400, 57600, 74880, 115200, 230400, 250000, 256000, 460800,
                 500000, 576000, 921600, 1000000, 2000000]

line_ending_list = ["None", "New line", "Carriage return", "NL + CR"]


def get_serial_ports_list():
    ports = serial.tools.list_ports.comports()
    com_names = []
    for port, desc, hwid in sorted(ports):
        com_names.append(port)
    if len(com_names) == 0:
        com_names.append("None")
    return com_names


def refresh_serials_port():
    _com_ports = get_serial_ports_list()
    selected_port.set(_com_ports[0])
    _menu = com_port_selector['menu']
    _menu.delete(0, "end")
    for string in _com_ports:
        _menu.add_command(label=string,
                          command=lambda value=string: selected_port.set(value))


def connect(port_name, baud_rate):
    ser.port = port_name
    ser.baudrate = baud_rate
    try:
        ser.open()
    except:
        messagebox.showerror(title="Serial Fatal Error", message="Serial connection failed!")
    finally:
        if ser.isOpen():
            connect_btn["state"] = "disable"
            disconnect_btn["state"] = "normal"
            com_port_selector["state"] = "disable"
            refresh_ser_ports_btn["state"] = "disable"
            baudrate_selector["state"] = "disable"


def disconnect():
    global send_buff
    send_buff = []
    ser.close()
    connect_btn["state"] = "normal"
    disconnect_btn["state"] = "disable"
    com_port_selector["state"] = "normal"
    refresh_ser_ports_btn["state"] = "normal"
    baudrate_selector["state"] = "normal"


def fail_disconnect():
    disconnect()
    messagebox.showerror(title="Connection broken", message="Serial connection broken!")


def on_closing():
    if ser.isOpen():
        ser.close()
    sys.exit()


def clear_data_log():
    if messagebox.askyesno("", "Czy na pewno chcesz wyczyścić log?", icon='question'):
        log_textbox.delete('1.0', END)


def save_to_file():
    datafile = asksaveasfile(initialfile='terminal.txt',
                             defaultextension=".txt",
                             filetypes=[("Text Documents", "*.txt")])
    if datafile is not None:
        datafile.write(log_textbox.get('1.0', END))
        datafile.close()


msg_new_line = True


def ser_send(text):
    global msg_new_line
    text = text
    if var_enable_echo.get():
        if not msg_new_line:
            log_textbox.insert(END, "\n")
        if var_enable_time_logging.get():
            datetime_now = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S.%f")
            log_textbox.insert(END, f"[{datetime_now}] ")
        if text[-1:] == "\n":
            log_textbox.insert(END, selected_port.get() + " <- " + text)
        else:
            log_textbox.insert(END, selected_port.get() + " <- " + text + "\n")
        log_textbox.see(END)
        msg_new_line = True
    try:
        if ser.isOpen():
            ser.write(text.encode('UTF-8'))

    except serial.SerialException:
        fail_disconnect()
    except:
        print("other exception")


lp = 1
plot_data = []
plot2_data = []
mm_last_mode = "OFF"
mm_last_unit = "OFF"
mm_last_mode2 = "OFF"
mm_last_unit2 = "OFF"


def clear_mm_data_log():
    global lp
    if messagebox.askyesno("", "Czy na pewno chcesz wyczyścić log?", icon='question'):
        log_mm_data.delete('1.0', END)
        log_mm_data.insert(END, 'lp,date and time,mode,value,unit\n')
        lp = 1


def save_mm_data_to_file():
    datafile = asksaveasfile(initialfile='multimeter_log.csv',
                             defaultextension=".csv",
                             filetypes=[("CSV Documents", "*.csv"), ("Text Documents", "*.txt")])
    if datafile is not None:
        datafile.write(log_mm_data.get('1.0', END))
        datafile.close()


def setup_plot(y_label="OFF (OFF)", y2_label=None):
    ax.set_xlabel("Samples")
    ax.set_ylabel(ylabel=y_label, color="#1685f5")
    ax.set_xlim(0, 200)
    ax.grid(color='#aaaaaa', linestyle='--', linewidth=1)
    ax.minorticks_on()
    ax.autoscale_view()

    if y2_label is not None:
        ax2.set_ylabel(ylabel=y2_label, color="#f00")

    else:
        ax2.set_ylabel(ylabel="OFF (OFF)", color="#f00")
    ax2.yaxis.tick_right()
    ax2.yaxis.set_label_position("right")
    ax2.minorticks_on()
    ax2.autoscale_view()


def plot_clear():
    global plot_data, mm_last_mode, mm_last_unit, plot2_data
    ax.cla()
    ax2.cla()
    setup_plot(f"{mm_last_mode} ({mm_last_unit})")
    plot_data = []
    plot2_data = []
    ax.plot(plot_data, color="blue")
    ax2.plot(plot2_data, color="blue")
    canvas.draw_idle()


def plot_save():
    plot_file = asksaveasfile(initialfile='multimeter_plot.png',
                             defaultextension=".png",
                             filetypes=[("PNG images", "*.png"), ("PDF file", "*.pdf")])
    if plot_file is not None:
        plot_file.close()
        fig.savefig(plot_file.name)


def ser_receive():
    global msg_new_line, lp, mm_last_mode, mm_last_unit, plot_data, mm_last_mode2, mm_last_unit2, plot2_data
    try:
        if ser.isOpen():
            if ser.in_waiting:
                msg = ser.read(ser.in_waiting).decode('UTF-8')
                # multimeter tab
                mm_data = msg.strip().split(' ')
                if len(mm_data) == 3:
                    mode_var.set(mm_data[0])
                    disp_var.set(mm_data[1] + mm_data[2])

                    log_mm_data.insert(END, str(lp) + ',' + datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S.%f") +
                                       ',' + mm_data[0] + ',' + mm_data[1].strip() + ',' + mm_data[2] + '\n')
                    log_mm_data.see("end")
                    lp += 1

                    if check_is_number(mm_data[1]):
                        if mm_last_mode != mm_data[0] or mm_last_mode2 != "OFF":
                            mm_last_mode = mm_data[0]
                            mm_last_unit = mm_data[2].strip()
                            mm_last_mode2 = "OFF"
                            mm_last_unit2 = "OFF"
                            ax.cla()
                            ax2.cla()
                            setup_plot(f"{mm_data[0]} ({mm_data[2].strip()})")
                            plot_data = []
                            plot2_data = []
                        if len(plot_data) > 200:
                            plot_data.pop(0)
                            ax.cla()
                            setup_plot(f"{mm_data[0]} ({mm_data[2].strip()})")
                        plot_data.append(float(mm_data[1]))
                        ax.plot(plot_data, color="#1685f5")
                        canvas.draw_idle()

                elif len(mm_data) == 6:
                    mode_var.set(mm_data[0])
                    disp_var.set(mm_data[1] + mm_data[2])

                    log_mm_data.insert(END, str(lp) + ',' + datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S.%f") +
                                       ',' + mm_data[0] + ',' + mm_data[1].strip() + ',' + mm_data[2] + '\n')
                    log_mm_data.see("end")
                    lp += 1

                    if check_is_number(mm_data[1]):
                        if mm_last_mode != mm_data[0] or mm_last_mode2 != mm_data[3]:
                            mm_last_mode = mm_data[0]
                            mm_last_unit = mm_data[2].strip()
                            mm_last_mode2 = mm_data[3]
                            mm_last_unit2 = mm_data[5].strip()
                            ax.cla()
                            setup_plot(f"{mm_data[0]} ({mm_data[2].strip()})",f"{mm_data[3]} ({mm_data[5].strip()})")
                            plot_data = []
                            plot2_data = []
                        if len(plot_data) > 200:
                            plot_data.pop(0)
                            plot2_data.pop(0)
                            ax.cla()
                            ax2.cla()
                            setup_plot(f"{mm_data[0]} ({mm_data[2].strip()})",f"{mm_data[3]} ({mm_data[5].strip()})")

                        plot_data.append(float(mm_data[1]))
                        plot2_data.append(float(mm_data[4]))

                        ax.plot(plot_data, color="#1685f5")
                        ax2.plot(plot2_data, color="#f00")

                        canvas.draw_idle()
                # end multimeter tab
                if msg_new_line:
                    if var_enable_time_logging.get():
                        datetime_now = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S.%f")
                        log_textbox.insert(END, f"[{datetime_now}] {selected_port.get()} -> {msg}")
                    else:
                        log_textbox.insert(END, f"{selected_port.get()} -> {msg}")
                else:
                    log_textbox.insert(END, f"{msg}")

                log_textbox.see(END)
                if msg[-1:] != '\n':
                    msg_new_line = False
                else:
                    msg_new_line = True

    except serial.SerialException:
        fail_disconnect()
    except Exception as e:
        print(f"other exception {e}")

    gui.after(10, ser_receive)


send_buff = []


def ser_send_thread():
    global send_buff
    while True:
        if len(send_buff) > 0:
            ser_send(send_buff.pop(0))
        time.sleep(0.001)


def add_to_send_buff(event):
    global send_buff
    if len(input_send.get()) > 0 and ser.isOpen():
        text = input_send.get()
        if selected_line_ending.get() == "New line":
            text += "\n"
        elif selected_line_ending.get() == "Carriage return":
            text += "\r"
        elif selected_line_ending.get() == "NL + CR":
            text += "\n\r"
        send_buff.append(text)
        input_send.delete(0, 'end')


def send_cmd(cmd):
    send_buff.append(cmd + "\n")


def auto_trig():
    if var_enable_auto_trigger.get() and ser.isOpen():
        send_cmd("MEASURE")
    gui.after(1000, lambda: auto_trig())


serial_send_thread = threading.Thread(target=ser_send_thread)
serial_send_thread.daemon = True
serial_send_thread.start()

gui = Tk()
gui.title("Serial term")
gui.geometry("900x590")
gui.resizable(False, False)
try:
    gui.iconbitmap(resource_path("icon.ico"))
except:
    pass

tabCtrl = ttk.Notebook(gui)
tab_terminal = ttk.Frame(tabCtrl)  # terminal tab
tab_mm = ttk.Frame(tabCtrl)  # multimeter tab
tab_plot = ttk.Frame(tabCtrl)  # multimeter tab

tabCtrl.add(tab_terminal, text='Terminal')
tabCtrl.add(tab_mm, text='Multimeter')
tabCtrl.add(tab_plot, text='Multimeter Plot')
tabCtrl.pack(expand=1, fill="both")


######### terminal tab #########
# com port
com_ports = get_serial_ports_list()

selected_port = StringVar(gui)
selected_port.set(com_ports[0])

com_port_selector = OptionMenu(tab_terminal, selected_port, *com_ports)
com_port_selector.place(x=5, y=24, width=150)
Label(tab_terminal, text="Port:", justify=LEFT).place(x=5, y=5)
refresh_ser_ports_btn = Button(tab_terminal, text="Refresh", command=lambda: refresh_serials_port())
refresh_ser_ports_btn.place(x=93, y=5, width=60, height=18)

# baudrate
selected_baudrate = StringVar(gui)
selected_baudrate.set("115200")
Label(tab_terminal, text="Baudrate:", justify=LEFT).place(x=160, y=5)
baudrate_selector = OptionMenu(tab_terminal, selected_baudrate, *baudrate_list)
baudrate_selector.place(x=160, y=24, width=150)

# connect btn
connect_btn = Button(tab_terminal, text="Connect", fg="#00dd00", command=lambda: connect(selected_port.get(),
                                                                                        selected_baudrate.get()))
connect_btn.place(x=170 + 150, y=5, width=80, height=48)

disconnect_btn = Button(tab_terminal, text="Disconnect", fg="#dd0000", command=lambda: disconnect())
disconnect_btn.place(x=260 + 150, y=5, width=80, height=48)
disconnect_btn["state"] = "disable"

# terminal log
frame_log = Frame(tab_terminal)
log_textbox_sb = Scrollbar(frame_log)
log_textbox_sb.pack(side=RIGHT, fill='y')
log_textbox = Text(frame_log, height=30, width=45, background="#111", foreground="#4287f5",
                   insertbackground="#ffffff", yscrollcommand=log_textbox_sb.set)
log_textbox_sb.config(command=log_textbox.yview)
log_textbox.place(x=0, y=0, width=885, height=450)
frame_log.place(x=0, y=55, width=900, height=450)

clear_log_btn = Button(tab_terminal, text="Clear log", fg="#777700", command=lambda: clear_data_log())
clear_log_btn.place(x=810, y=5, width=85, height=20)
save_log_btn = Button(tab_terminal, text="Save log", fg="#1f75de", command=lambda: save_to_file())
save_log_btn.place(x=810, y=30, width=85, height=20)

selected_line_ending = StringVar(gui)
selected_line_ending.set("New line")

Label(tab_terminal, text="Line ending:", justify=LEFT).place(x=650, y=3)
line_ending_selector = OptionMenu(tab_terminal, selected_line_ending, *line_ending_list)
line_ending_selector.place(x=650, y=23, width=150)

# send
input_send = Entry(tab_terminal)
input_send.place(x=5, y=510, width=805)
input_send.bind('<Return>', add_to_send_buff)
Button(tab_terminal, text="Send", fg="#000", command=lambda: add_to_send_buff(None)).place(x=815, y=510,
                                                                                          width=80, height=20)

var_enable_echo = IntVar()
var_enable_echo.set(1)
Checkbutton(tab_terminal, text='Enable local echo', variable=var_enable_echo, onvalue=1, offvalue=0).place(x=0, y=532)

var_enable_time_logging = IntVar()
var_enable_time_logging.set(1)
Checkbutton(tab_terminal, text='Enable time logging', variable=var_enable_time_logging,
            onvalue=1, offvalue=0).place(x=170, y=532)


######### multimeter tab #########
mode_var = StringVar(gui)
mode_var.set("OFF")
disp_var = StringVar(gui)
disp_var.set("OFF")

Label(tab_mm, text="Mode:", justify=LEFT, width=5).place(x=10, y=9)
multimeter_mode = Label(tab_mm, textvariable=mode_var, justify=CENTER, width=5, bd=3, relief="sunken", padx=5)
multimeter_mode.config(font=('Roboto Mono', 28, 'bold'), background='#dddddd', fg="#004f46")
multimeter_mode.place(x=10, y=29)

multimeter_disp = Label(tab_mm, textvariable=disp_var, justify=LEFT, width=10, bd=5, relief="sunken", padx=15)
multimeter_disp.config(font=('Roboto Mono', 40, 'bold'), background='#dddddd')
multimeter_disp.place(x=140, y=5)

Button(tab_mm, text="MEASURE", fg="#ed1202", command=lambda: send_cmd("MEASURE")).place(x=510, y=5, width=100)

var_enable_auto_trigger = IntVar()
var_enable_auto_trigger.set(0)
Checkbutton(tab_mm, text='Auto Trigger', variable=var_enable_auto_trigger,
            onvalue=1, offvalue=0).place(x=510, y=40)
auto_trig()

Button(tab_mm, text="BACKLIGHT", fg="#4287f5", command=lambda: send_cmd("BACKLIGHT")).place(x=795, y=5, width=100)



frame_log_mm_data = Frame(tab_mm)
log_mm_data_sb = Scrollbar(frame_log_mm_data)
log_mm_data_sb.pack(side=RIGHT, fill='y')
log_mm_data = Text(frame_log_mm_data, height=30, width=45, background="#111", foreground="#ffc94d",
                   insertbackground="#ffffff", yscrollcommand=log_mm_data_sb.set)
log_mm_data_sb.config(command=log_mm_data.yview)
log_mm_data.place(x=0, y=0, width=885, height=450)
frame_log_mm_data.place(x=0, y=110, width=900, height=450)
log_mm_data.insert(END, 'lp,date and time,mode,value,unit\n')

clear_mm_log_btn = Button(tab_mm, text="Clear", fg="#777700", command=lambda: clear_mm_data_log())
clear_mm_log_btn.place(x=740, y=75, width=75)
save_mm_log_btn = Button(tab_mm, text="Save", fg="#1f75de", command=lambda: save_mm_data_to_file())
save_mm_log_btn.place(x=820, y=75, width=75)


######### multimeter plot tab #########
plotting_frame = Frame(tab_plot)
fig = Figure()
fig.suptitle("Pomiary")
ax = fig.add_subplot()
ax2 = ax.twinx()
setup_plot()

canvas = FigureCanvasTkAgg(fig, master=plotting_frame)
canvas.get_tk_widget().place(x=0, y=0, width=900, height=560)
canvas.draw()
plotting_frame.place(x=0, y=0, width=900, height=560)

clear_mm_plot_btn = Button(tab_plot, text="Save", fg="#00cc00", command=lambda: plot_save())
clear_mm_plot_btn.place(x=820, y=5, width=75, height=20)

clear_mm_plot_btn = Button(tab_plot, text="Clear", fg="#cc0000", command=lambda: plot_clear())
clear_mm_plot_btn.place(x=820, y=28, width=75, height=20)


gui.after(1, ser_receive)
gui.protocol("WM_DELETE_WINDOW", on_closing)
gui.mainloop()

sys.exit()
