from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import tkinter.ttk as ttk
from hashlib import md5, sha1
import os
import openpyxl
import time
import threading

entry_points = []
choosen_algo = [False, False]
excel_filename = ["", ""]
wasnt_counted_dirs = []
wasnt_counted_files = []
WORK = False
EXIT = False


def deploy_files():
    global entry_points
    temp_points = [i for i in entry_points]
    files_to_delete = []
    for dir in temp_points:
        for file in temp_points:
            if (dir in file) and (dir != file):
                files_to_delete.append(file)
    for file in files_to_delete:
        try:
            temp_points.remove(file)
        except ValueError:
            continue
    all_files = []
    for file in temp_points:
        if os.path.isfile(file):
            all_files.append(file)
        else:
            for addr, dir, filenames in os.walk(file):
                for filename in filenames:
                    all_files.append(addr + '\\' + filename.replace('/', '\\'))
    return all_files


def crypt(path):
    try:
        file_stream_in = open(path, 'rb')
        file_data = ""
        file_data = file_stream_in.read()
        file_stream_in.close()
        index = 0
        cell = []
        crypted_data = []
        try:
            index = path.rindex('\\')
        except ValueError:
            pass
        file_name = path[index + 1: len(path): 1]
        if len(file_data) != 0:
            if choosen_algo == [False, True]:
                crypted_data.append(sha1(file_data).hexdigest())
            if choosen_algo == [True, False]:
                crypted_data.append(md5(file_data).hexdigest())
            if choosen_algo == [True, True]:
                crypted_data.append(md5(file_data).hexdigest())
                crypted_data.append(sha1(file_data).hexdigest())
            cell.append(file_name)
            for i in crypted_data:
                cell.append(i)
            cell.append(path)
            return cell
        else:
            if choosen_algo == [True, True]:
                return [file_name, 'Файл пуст', '', path]
            else:
                return [file_name, 'Файл пуст', path]
    except PermissionError:
        wasnt_counted_files.append(path)
        pass


def init_excel(sheet):
    sheet["A1"].value = "ИМЯ ФАЙЛА"
    if choosen_algo == [True, False]:
        sheet["B1"].value = "MD5 СУММА"
        sheet["C1"].value = "ПУТЬ К ФАЙЛУ"
    if choosen_algo == [False, True]:
        sheet["B1"].value = "SHA1 СУММА"
        sheet["C1"].value = "ПУТЬ К ФАЙЛУ"
    if choosen_algo == [True, True]:
        sheet["B1"].value = "MD5 СУММА"
        sheet["C1"].value = "SHA1 СУММА"
        sheet["D1"].value = "ПУСТЬ К ФАЙЛУ"


def count_checksum():
    global WORK
    global Progressbar
    global EXIT
    while 1:
        time.sleep(0.5)
        if WORK:
            time_start = time.time()
            iter = 2
            excel_path = excel_filename[0] + '\\' + excel_filename[1] + ".xlsx"
            try:
                os.remove(excel_path)
                if excel_path in entry_points:
                    entry_points.remove(excel_path)
            except FileNotFoundError:
                pass
            excel_stream_out = openpyxl.Workbook()
            sheet = excel_stream_out["Sheet"]
            init_excel(sheet)
            files = deploy_files()
            excel_output = []
            pb_all_count = len(files) * 2
            pb_counter = 100 / pb_all_count
            for file in files:
                Progressbar['value'] += pb_counter
                excel_output.append(crypt(file))
            for i in excel_output:
                Progressbar['value'] += pb_counter
                for j in range(0, len(i)):
                    sheet[chr(ord("A") + j) + str(iter)].value = i[j]
                iter += 1
            excel_stream_out.save(excel_path)
            messagebox.showinfo("OK",
                                "Работа Выполнена\nБыло затрачено: {} сек".format(round(time.time() - time_start, 2)))
            excel_output.clear()
            Progressbar['value'] = 0
            Progressbar.place(x=10, y=10, width=0)
            WORK = False
        if EXIT:
            exit()


def choose_dir():
    try:
        directory = filedialog.askdirectory()
        directory = os.path.abspath(directory)
        if directory not in entry_points:
            list_box.insert(END, directory)
            entry_points.append(directory)
    except PermissionError:
        messagebox.showerror("Permission Error", "Нет прав доступа к выбранной директорий")
        pass


def choose_files():
    try:
        choosen_files = filedialog.askopenfilenames(filetypes=(("All Files", "*.*"), ("Text Files", "*.txt")))
        for file in choosen_files:
            file = os.path.abspath(file)
            if file not in entry_points:
                list_box.insert(END, file)
                entry_points.append(file)
    except PermissionError:
        messagebox.showerror("Permission Error", "Нет прав доступа к выбранному файлу")
        pass


def delete_insertion():
    global entry_points
    delete_list = list_box.curselection()
    to_delete = []
    for i in range(0, len(delete_list)):
        to_delete.append(entry_points[delete_list[i]])
    delete_count = 0
    for element in entry_points:
        if element in to_delete:
            ind = entry_points.index(element, 0, len(entry_points)) - delete_count
            list_box.delete(ind, ind)
            delete_count += 1
    entry_points = [item for item in entry_points if item not in to_delete]


def md5_choose():
    choosen_algo[0] = not choosen_algo[0]


def sha1_choose():
    choosen_algo[1] = not choosen_algo[1]


def choose_excel_directory():
    try:
        directory = os.path.abspath(filedialog.askdirectory())
        excel_filename[0] = directory
        excel_directory_text.configure(state=NORMAL)
        excel_directory_text.delete(1.0, END)
        excel_directory_text.insert(END, directory)
        excel_directory_text.configure(state=DISABLED)
    except PermissionError:
        messagebox.showerror("Permission Error", "Нет прав доступа к выбранному файлу")
        pass


def ok_click():
    if excel_entry["state"] == NORMAL:
        excel_filename[1] = excel_entry.get()
        if excel_filename[1] != "":
            excel_entry["state"] = DISABLED

    else:
        excel_entry["state"] = NORMAL
        if len(excel_entry.get()) != 0:
            excel_entry.delete(0, END)


def do_work():
    global WORK
    if len(excel_filename[0]) != 0 and len(excel_filename[1]) != 0 and choosen_algo != [False, False] and len(
            entry_points) != 0:
        WORK = True
        Progressbar.place(x=125, y=400, width=400)
    else:
        messagebox.showerror("Incomplete form", "Пожалуйста, заполните все поля")


def resource_path(relative):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative)
    else:
        return os.path.join(os.path.abspath("."), relative)


thread1 = threading.Thread(target=count_checksum)
thread1.start()
window = Tk()
window.title("Counter v:0.1.0")
window.iconbitmap(resource_path("icon.ico"))
window.geometry('650x500')
window.maxsize(650, 500)
window.minsize(650, 500)
choose_dir_btn = Button(window, text="Выбрать директорию", command=choose_dir, font=("Comic Sans", 10))
choose_file_btn = Button(window, text="Выбрать файлы", command=choose_files, font=("Comic Sans", 10))
delete_btn = Button(window, text="X", command=delete_insertion)
scroll_bar = Scrollbar(window, elementborderwidth=100)
list_box = Listbox(window, state=NORMAL, selectmode="multiple", width=75, height=10, yscrollcommand=scroll_bar.set)
scroll_bar.configure(command=list_box.yview)
Algo_label = Label(window, text="Алгоритмы:", font=("Comis Sans", 16))
md5_checkbutton = Checkbutton(window, text="MD5", command=md5_choose)
sha1_checkbutton = Checkbutton(window, text="SHA1", command=sha1_choose)
choose_dir_excel_btn = Button(window, text="Выбрать директорию", command=choose_excel_directory,
                              font=("Comic Sans", 10))
scroll_bar_excel = Scrollbar(window, orient=HORIZONTAL)
excel_directory_text = Text(window, height=1, width=60)
scroll_bar_excel.configure(command=excel_directory_text.yview)
excel_directory_text.configure(yscrollcommand=scroll_bar_excel.set, state=DISABLED)
excel_label = Label(window, text="Введите название:", font=("Comis Sans", 12))
excel_entry = Entry(window, width=40, state=NORMAL)
excel_ok_btn = Button(window, text="OK", font=("Comic Sans", 10), command=ok_click)
Start_btn = Button(window, text="Посчитать", font=("Comic Sans", 12), command=do_work)
Progressbar = ttk.Progressbar(window, mode="determinate")
Creators_label = Label(window, text="Created by: Igor Pavlov, Svetlana Bogorodskaya, Vladislav Bermishev",
                       foreground="dim gray", font=("Cambria", 9))
list_box.place(x=10, y=10)
scroll_bar.place(x=465, y=10, height=165)
choose_dir_btn.place(x=490, y=10)
choose_file_btn.place(x=490, y=50)
delete_btn.place(x=490, y=90)
Algo_label.place(x=10, y=180)
md5_checkbutton.place(x=120, y=185)
sha1_checkbutton.place(x=170, y=185)
choose_dir_excel_btn.place(x=10, y=225)
excel_directory_text.place(x=150, y=230)
scroll_bar_excel.place(x=150, y=250, width=485)
excel_label.place(x=10, y=280)
excel_entry.place(x=150, y=282)
excel_ok_btn.place(x=395, y=280)
Start_btn.place(x=280, y=350, width=90)
Creators_label.place(x=0, y=478)
window.mainloop()
EXIT = True
