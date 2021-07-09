import tkinter as tk
from urllib.request import urlopen
from urllib.error import HTTPError
import json


def getHTML(url):
    try:
        html = urlopen(url)
    except HTTPError as e:
        return None
    return html


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.warning=tk.Label(self)
        # self.warning["text"]="warning"
        self.warning.place(x=50,y=0)
        self.warning.pack()
        self.l_bin=tk.Label(self,text="bin: ")
        self.l_bin.place(x=0,y=20)
        self.l_bin.pack()
        self.bin = tk.Entry(self)
        self.bin.place(x=50,y=20)
        self.bin.pack()
        self.json = tk.Button(self)
        self.json["text"]="Download"
        self.json["command"]=self.download
        self.json.pack(side="bottom")

    def download(self):
        bin_=self.bin.get()
        html = getHTML(f"https://stat.gov.kz/api/juridical/counter/api/?bin={bin_}&lang=ru")
        if html == None:
            self.warning["text"]="HTML could not be found"
        else:
            self.warning["text"] = "JSON was downloaded"
            prepare = html.read().decode()
            data = json.loads(prepare)
            filename=data["obj"]["name"]
            with open(f'{filename}.json', 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)


root = tk.Tk()
app = Application(master=root)
app.mainloop()

#https://stat.gov.kz/api/juridical/counter/api/?bin={BIN}&lang={LANG}
