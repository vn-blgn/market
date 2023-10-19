from tkinter import *
from src import handlers


def main():
    root = Tk()
    root.title("MARKET")
    root.iconbitmap(default="src/favicon.ico")
    root.geometry("1000x500")

    root.rowconfigure(0, minsize=800, weight=1)
    root.columnconfigure(1, minsize=800, weight=1)

    txt_edit = Text(root)
    fr_buttons = Frame(root)
    btn_open = Button(fr_buttons, text="Обновить котировки", command=lambda: handlers.update_quotes(txt_edit, END))
    btn_save = Button(fr_buttons, text="Загрузить файл", command=lambda: handlers.upload_file(txt_edit, END))

    btn_open.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
    btn_save.grid(row=1, column=0, sticky="ew", padx=5)

    fr_buttons.grid(row=0, column=0, sticky="ns")
    txt_edit.grid(row=0, column=1, sticky="nsew")

    root.mainloop()


if __name__ == '__main__':
    main()
