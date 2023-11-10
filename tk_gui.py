from tkinter import Tk, Button
from tkinter.messagebox import showinfo
from tkinter.filedialog import askopenfilename

from parser_json import Parser


class Gui:

    def __init__(self, master):
        self.master = master
        self.master.title('JSON Parser')
        self.master.geometry('240x240')
        Button(master, text='Загрузить JSON', command=self.upload).\
            grid(row=0, column=0, padx=60, pady=60)

    def upload(self):
        file = askopenfilename(filetypes=[("Json files", ".json")])
        Parser(file)

        showinfo(title='Окончание операции', message='Конвертация завершена')

        self.master.destroy()


if __name__ == '__main__':
    root = Tk()
    mw = Gui(root)
    root.mainloop()
