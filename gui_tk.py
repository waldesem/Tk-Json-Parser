from tkinter import Button, Tk, filedialog

from json_parser import JsonFile, ExcelFile

class MainWindow:

    def __init__(self, master, title, geometry):
        self.master = master
        self.title = title
        self.geometry = geometry
        self.master.title(title)
        self.master.geometry(geometry)
        Button(master, text='Загрузить JSON', command=self.upload).\
            grid(row=0, column=0, padx=60, pady=60)

    def close(self):
        self.master.destroy()

    def upload(self):
        file = filedialog.askopenfilename(filetypes=[("Json files", ".json")])
        print(file)
        data = JsonFile(file)
        export = ExcelFile('anketa.xlsx')
        export.upload_data(data)
        export.wb.save(file.replace('json', 'xlsx'))

        self.close()

if __name__ == '__main__':
    root = Tk()
    mw = MainWindow(root,'JSON Parser', '240x240')
    root.mainloop()