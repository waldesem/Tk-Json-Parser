from tkinter import Tk, Button
from tkinter.messagebox import showinfo
from tkinter.filedialog import askopenfilename

from json_parser import convert

class Gui:

    def __init__(self, master):
        self.master = master
        self.master.title('JSON Parser')
        self.master.geometry('240x240')
        Button(master, text='Загрузить JSON', command=self.upload).\
            grid(row=0, column=0, padx=60, pady=60)

    def upload(self):
        file = askopenfilename(filetypes=[("Json files", ".json")])

        convert(file)
        
        showinfo(title='Окончание операции', message='Конвертация завершена')

        self.master.destroy()


if __name__ == '__main__':
    root = Tk()
    Gui(root)
    root.mainloop()
