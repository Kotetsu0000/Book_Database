import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox
from tkinter import filedialog
import pandas as pd
import requests
import xml.etree.ElementTree as et
import json

class Application(tk.Frame):
    def __init__(self, master=None):
        tk.Frame.__init__(self, master)
        self.pack(expand=1, fill=tk.BOTH, anchor=tk.NW)
        self.master = master
        self.camera_list = None
        self.camera_window = None
        self.get_camera = None
        self.set_data()
        self.create_widgets()

    def set_data(self):
        self.colname_list = ['タイトル', '著者', '出版社', '件名標目','保管場所']  # 検索結果に表示させる列名
        self.in_colname_list = ['isbn','タイトル', '著者', '出版社', '件名標目','保管場所']  # 検索結果に表示させる列名
        try:
            self.data = pd.read_csv('book_data.csv')
        except:
            self.data = pd.DataFrame(None, columns=self.in_colname_list)
            self.data.to_csv('book_data.csv', encoding="utf-8", index=False)
        self.data = self.data.reset_index(drop=True)
        
        self.width_list = [200, 100, 100, 100, 50]

    def create_widgets(self):
        self.main_window = tk.PanedWindow(self, orient="vertical")
        self.main_window.pack(expand=True, fill=tk.BOTH, side=tk.TOP)

        search_region_frame = ttk.Frame(self.main_window, )
        search_region_frame.pack(fill=tk.X, side="top")
        self.button_region = ttk.PanedWindow(search_region_frame, orient="horizontal")
        self.button_region.pack(fill=tk.Y, side="right")
        self.search_region = ttk.PanedWindow(search_region_frame, orient="horizontal")
        self.search_region.pack(expand=True,fill=tk.BOTH, side="left")

        self.search_num_region = ttk.PanedWindow(self.main_window, orient="horizontal")
        self.search_num_region.pack(fill=tk.X, side="top")

        self.table_region = ttk.PanedWindow(self.main_window, orient="vertical")
        self.table_region.pack(expand=True,fill=tk.BOTH, side="left")

        self.create_search_widgets()
        self.create_table_widgets()

    def create_search_widgets(self):
        search_frame = tk.Frame(self.search_region)
        self.search_region.add(search_frame)

        name_frame = tk.Frame(search_frame)
        name_frame.pack(fill=tk.Y, side="left")

        in_frame = tk.Frame(search_frame)
        in_frame.pack(expand=True,fill=tk.BOTH, side="left")

        self.label_title = ttk.Label(name_frame, text='タイトル')
        self.label_title.pack(side="top",pady=3)
        self.search_str_title = tk.StringVar()
        self.search_box_title = ttk.Entry(in_frame, justify="left",textvariable=self.search_str_title)
        self.search_box_title.pack(fill=tk.X, side="top", padx=2, pady=2)
        self.search_box_title.bind("<Return>", self.search_table)

        self.label_auther = ttk.Label(name_frame, text='著者')
        self.label_auther.pack(side="top",pady=3)
        self.search_str_auther = tk.StringVar()
        self.search_box_auther = ttk.Entry(in_frame, justify="left",textvariable=self.search_str_auther)
        self.search_box_auther.pack(fill=tk.X, side="top", padx=2, pady=2)
        self.search_box_auther.bind("<Return>", self.search_table)

        self.label_publisher = ttk.Label(name_frame, text='出版社')
        self.label_publisher.pack(side="top",pady=3)
        self.search_str_publisher = tk.StringVar()
        self.search_box_publisher = ttk.Entry(in_frame, justify="left",textvariable=self.search_str_publisher)
        self.search_box_publisher.pack(fill=tk.X, side="top", padx=2, pady=2)
        self.search_box_publisher.bind("<Return>", self.search_table)

        self.label_subject = ttk.Label(name_frame, text='件名標目')
        self.label_subject.pack(side="top",pady=3)
        self.search_str_subject = tk.StringVar()
        self.search_box_subject = ttk.Entry(in_frame, justify="left",textvariable=self.search_str_subject)
        self.search_box_subject.pack(fill=tk.X, side="top", padx=2, pady=2)
        self.search_box_subject.bind("<Return>", self.search_table)

        self.label_place = ttk.Label(name_frame, text='保管場所')
        self.label_place.pack(side="top",pady=3)
        self.search_str_place = tk.StringVar()
        self.search_box_place = ttk.Entry(in_frame, justify="left",textvariable=self.search_str_place)
        self.search_box_place.pack(fill=tk.X, side="top", padx=2, pady=2)
        self.search_box_place.bind("<Return>", self.search_table)

        self.button = tk.Button(self.button_region, text=u'ISBNを入力', width=20, height=1, command=self.button_pushed)
        self.button.pack(side = tk.TOP, pady = 3, padx = 5)

        self.button = tk.Button(self.button_region, text=u'CSVで保存', width=20, height=1, command=self.save_csv)
        self.button.pack(side = tk.TOP, pady = 3, padx = 5)

        self.button = tk.Button(self.button_region, text=u'更新', width=20, height=1, command=self.search_table)
        self.button.pack(side = tk.BOTTOM, pady = 3, padx = 5)

        self.button = tk.Button(self.button_region, text=u'選択を削除', width=20, height=1, command=self.del_data)
        self.button.pack(side = tk.BOTTOM, pady = 3, padx = 5)

    def search_table(self, event=None):
        bool_data = [True]*len(self.data)
        if self.search_str_title.get()!='':
            bool_data = bool_data & self.data['タイトル'].str.contains(self.search_str_title.get(), na=False)
        if self.search_str_auther.get()!='':
            bool_data = bool_data & self.data['著者'].str.contains(self.search_str_auther.get(), na=False)
        if self.search_str_publisher.get()!='':
            bool_data = bool_data & self.data['出版社'].str.contains(self.search_str_publisher.get(), na=False)
        if self.search_str_subject.get()!='':
            bool_data = bool_data & self.data['件名標目'].str.contains(self.search_str_subject.get(), na=False)
        if self.search_str_place.get()!='':
            bool_data = bool_data & self.data['件名標目'].str.contains(self.search_str_place.get(), na=False)
        result = self.data[bool_data]
        self.update_table_by_search_result(result)

    def update_table_by_search_result(self, result):
        self.table.delete(*self.table.get_children())
        self.result_text.set(f"検索結果：{len(result)}")
        for _, row in result.iterrows():
            self.table.insert("", "end", values=row[self.colname_list].to_list())

    def create_table_widgets(self):
        self.result_text = tk.StringVar()
        len_result = ttk.Label(self.search_num_region, textvariable=self.result_text)
        self.search_num_region.add(len_result)
        
        table_frame = tk.Frame(self.table_region)
        table_frame.pack(expand=True,fill=tk.BOTH, side="left")
        self.table = ttk.Treeview(table_frame)
        self.table["column"] = self.colname_list
        self.table["show"] = "headings"
        for i, (colname, width) in enumerate(zip(self.colname_list, self.width_list)):
            self.table.heading(i, text=colname)
            self.table.column(i, minwidth=width, width=width, stretch=True)
        self.table.grid(row=0, column=0, sticky='nsew')
        self.table.bind("<Double-1>", self.onDuble)

        ysb = tk.Scrollbar(table_frame, orient=tk.VERTICAL, width=16, command=self.table.yview)
        self.table.configure(yscrollcommand=ysb.set)
        ysb.grid(row=0, column=1, sticky='nsew')

        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        self.update_table_by_search_result(self.data)

    def button_pushed(self):
        root = tk.Toplevel()
        w = root.winfo_screenwidth()    #モニター横幅取得
        h = root.winfo_screenheight()   #モニター縦幅取得
        windw_w = 300
        windw_h = 75
        root.title("ISBNを入力")
        root.geometry(str(windw_w)+"x"+str(windw_h)+"+"+str((w-windw_w)//2)+"+"+str((h-windw_h)//2))
        #root.geometry("480x300")
        root.protocol("WM_DELETE_WINDOW", self.delete_window)
        root.grab_set()
        root.resizable(width=False, height=False)
        self.input_isbn_window = input_isbn(master=root, data=self.data)
        self.after(1000,self.update_sub_window)

    def wait_check_camera(self):
        if self.get_camera.camera_list == None:
            self.after(1000, self.wait_check_camera)
        else:
            self.camera_list = self.get_camera.camera_list
            self.after(1000, self.button_pushed)

    def no_mvoe(self):
        pass
    
    def delete_window(self):
        self.data = self.input_isbn_window.data
        self.search_table()
        self.input_isbn_window.master.destroy()

    def delete_window_2(self):
        self.data = self.input_data_window.data
        self.search_table()
        self.input_data_window.master.destroy()

    def update_sub_window(self):
        self.search_table()
        if self.input_isbn_window.master.winfo_exists():
            self.after(1,self.update_sub_window)

    def del_data(self):
        if len(self.table.selection())!=0:
            ret_info = messagebox.askyesno('削除確認', '選択された項目を削除しますか？')
            if ret_info:
                del_list = []
                for item in self.table.selection():
                    del_list.append(self.table.item(item)["values"])
                for del_item in del_list:
                    del_index = self.get_index(del_item)
                    self.data = self.data.drop(del_index)
                    self.data = self.data.reset_index(drop=True)
                self.search_table()
                #self.data.to_csv("./csv/book_data.csv", encoding="utf-8", index=False)
                self.data.to_csv('book_data.csv', encoding="utf-8", index=False)

    def save_csv(self, event=None):
        filename = filedialog.asksaveasfilename(
            title = "名前を付けて保存",
            filetypes = [("CSV", ".csv") ], # ファイルフィルタ
            initialdir = "./", # 自分自身のディレクトリ
            defaultextension = "csv"
            )
        self.data.to_csv(filename, encoding="utf-8", index=False)

    def onDuble(self, event):
        if len(self.table.selection())!=1:
            for item in self.table.selection():
                print(self.get_index(self.table.item(item)["values"]))
                print(self.table.item(item)["values"])
        else:
            root = tk.Toplevel()
            w = root.winfo_screenwidth()    #モニター横幅取得
            h = root.winfo_screenheight()   #モニター縦幅取得
            windw_w = 300
            windw_h = 225
            root.title("項目の編集")
            root.geometry(str(windw_w)+"x"+str(windw_h)+"+"+str((w-windw_w)//2)+"+"+str((h-windw_h)//2))
            root.protocol("WM_DELETE_WINDOW", self.no_mvoe)
            root.grab_set()
            root.resizable(width=False, height=False)
            self.input_data_window = change_data(master=root, data=self.data, index=self.get_index(self.table.item(self.table.selection()[0])["values"]), del_func=self.delete_window_2)
            
    def get_index(self,table_list):
        title, creator, publisher, subject, Place = table_list
        return list(self.data[((self.data['タイトル']==title)|(self.data['タイトル'].isna())) & ((self.data['著者']==creator) | (self.data['著者'].isna())) & ((self.data['出版社']==publisher)|(self.data['出版社'].isna())) & ((self.data['件名標目']==subject)|(self.data['件名標目'].isna()))].index)[0]

    def change_table_element(self, index, table_list):
        self.data.iloc[index]=table_list
        self.search_table()

class change_data(tk.Frame):
    def __init__(self, master=None, data=None, index=None, del_func=None):
        tk.Frame.__init__(self, master)
        self.pack(expand=1, fill=tk.BOTH, anchor=tk.NW)
        self.master = master
        self.data = data
        self.index = index
        self.del_func = del_func
        self.create_widgets()

    def create_widgets(self):
        self.label_isbn = tk.Label(self, text=u'ISBN')
        self.label_isbn.grid(row=0, column=0)
        self.str_isbn = tk.StringVar()
        self.str_isbn.set(self.data.iat[self.index,0])
        self.input_box_isbn = ttk.Entry(self, justify="left",textvariable=self.str_isbn)
        self.input_box_isbn.grid(row=0, column=1, pady = 5, padx = 5, sticky=tk.EW, columnspan=2)

        self.label_title = tk.Label(self, text=u'タイトル')
        self.label_title.grid(row=1, column=0)
        self.str_title = tk.StringVar()
        self.str_title.set(self.data.iat[self.index,1])
        self.input_box_title = ttk.Entry(self, justify="left",textvariable=self.str_title)
        self.input_box_title.grid(row=1, column=1, pady = 5, padx = 5, sticky=tk.EW, columnspan=2)

        self.label_creator = tk.Label(self, text=u'著者')
        self.label_creator.grid(row=2, column=0)
        self.str_creator = tk.StringVar()
        self.str_creator.set(self.data.iat[self.index,2])
        self.input_box_creator = ttk.Entry(self, justify="left",textvariable=self.str_creator)
        self.input_box_creator.grid(row=2, column=1, pady = 5, padx = 5, sticky=tk.EW, columnspan=2)

        self.label_publisher = tk.Label(self, text=u'出版社')
        self.label_publisher.grid(row=3, column=0)
        self.str_publisher = tk.StringVar()
        self.str_publisher.set(self.data.iat[self.index,3])
        self.input_box_publisher = ttk.Entry(self, justify="left",textvariable=self.str_publisher)
        self.input_box_publisher.grid(row=3, column=1, pady = 5, padx = 5, sticky=tk.EW, columnspan=2)

        self.label_subject = tk.Label(self, text=u'件名標目')
        self.label_subject.grid(row=4, column=0)
        self.str_subject = tk.StringVar()
        self.str_subject.set(self.data.iat[self.index,4])
        self.input_box_subject = ttk.Entry(self, justify="left",textvariable=self.str_subject)
        self.input_box_subject.grid(row=4, column=1, pady = 5, padx = 5, sticky=tk.EW, columnspan=2)

        self.label_place = tk.Label(self, text=u'保管場所')
        self.label_place.grid(row=5, column=0)
        self.str_place = tk.StringVar()
        self.str_place.set(self.data.iat[self.index,5])
        self.input_box_place = ttk.Entry(self, justify="left",textvariable=self.str_place)
        self.input_box_place.grid(row=5, column=1, pady = 5, padx = 5, sticky=tk.EW, columnspan=2)

        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=1)

        self.button = tk.Button(self, text=u'OK', command=self.change_table_element)
        self.button.grid(row=6, column=1, pady = 5, padx = 5, sticky=tk.EW)

        self.button = tk.Button(self, text=u'キャンセル', command=self.del_func)
        self.button.grid(row=6, column=2, pady = 5, padx = 5, sticky=tk.EW)

    def change_table_element(self):
        table_list = [self.str_isbn.get(), self.str_title.get(), self.str_creator.get(), self.str_publisher.get(), self.str_subject.get(), self.str_place.get()]
        self.data.iloc[self.index]=table_list
        self.data.to_csv('book_data.csv', encoding="utf-8", index=False)
        self.del_func()

class input_isbn(tk.Frame):
    def __init__(self, master=None, data=None):
        tk.Frame.__init__(self, master)
        self.pack(expand=1, fill=tk.BOTH, anchor=tk.NW)
        self.master = master
        self.data = data
        self.create_widgets()

    def create_widgets(self):
        self.label = tk.Label(self, text=u'追加したい本のISBNを入力してください。')
        self.label.pack(side=tk.TOP,anchor = tk.W)
        self.search_str_isbn = tk.StringVar()
        self.search_box_isbn = ttk.Entry(self, justify="left",textvariable=self.search_str_isbn)
        self.search_box_isbn.pack(fill=tk.X, side=tk.TOP, padx=2, pady=2)
        self.search_box_isbn.bind("<Return>", self.search)
        self.button = tk.Button(self, text=u'検索', command=self.search)
        self.button.pack(side = tk.BOTTOM, pady = 5, padx = 5)

    def search(self, event=None):
        if self.search_str_isbn.get()=='':
            messagebox.showerror('登録失敗', '入力されていません。')
        elif len(self.data[self.data['isbn'].astype(str)==self.search_str_isbn.get()])==0:
            isbn, title, creator, publisher, subject = self.fetch_book_data(self.search_str_isbn.get())
            if title!='':
                ret_info = messagebox.askyesno('追加確認', '追加しますか？\nISBN : '+isbn+'\nタイトル : '+title+'\n著者 : '+creator+'\n出版社 : '+publisher+'\n件名標目 : '+subject)
                if ret_info:
                    self.data = self.data.append({'isbn': isbn, 'タイトル': title, '著者': creator, '出版社': publisher, '件名標目': subject}, ignore_index=True)
                    self.data.to_csv('book_data.csv', encoding="utf-8", index=False)
                    self.search_str_isbn.set('')
            else:
                messagebox.showerror('登録失敗', '入力されたISBNコードと一致する本が存在ません。')
        else:
            messagebox.showerror('登録失敗', 'この本は既に登録されています。')
    
    def list_to_text(self, in_list):
        return_text = ''
        for i in in_list:
            return_text += str(i) + ', '
        return return_text[:-2]

    def xml_to_list_text(self, xml_data, key):
        ns = {'dc': 'http://purl.org/dc/elements/1.1/'}
        try:
            return_list = []
            for i in xml_data.findall(key, ns):
                if not(i.text in return_list):
                    return_list.append(i.text)
            return self.list_to_text(return_list)
        except: return ''

    def fetch_author_data_openlibrary(self, key):
        endpoint = 'https://openlibrary.org' + str(key) + '.json'
        res = requests.get(endpoint)
        res_json = json.loads(res.text)
        return res_json['name']

    def fetch_book_data_openlibrary(self, isbn):
        endpoint = 'https://openlibrary.org/isbn/' + str(isbn) + '.json'
        res = requests.get(endpoint)
        res_json = json.loads(res.text)
        author_list = []
        for key in res_json['authors']:
            author_list.append(self.fetch_author_data_openlibrary(key['key']))
        return isbn, res_json['title'], self.list_to_text(author_list), self.list_to_text(res_json['publishers']), self.list_to_text(res_json['subjects'])

    def fetch_book_data_NDL(self, isbn):
        endpoint = 'https://iss.ndl.go.jp/api/sru'
        params = {'operation': 'searchRetrieve',
                'query': f'isbn="{isbn}"',
                'recordPacking': 'xml'}
        res = requests.get(endpoint, params=params)
        root = et.fromstring(res.text)
        ns = {'dc': 'http://purl.org/dc/elements/1.1/'}
        try:
            title = root.find('.//dc:title', ns).text
        except:
            title = ''
        creator = self.xml_to_list_text(root, './/dc:creator')
        publisher = self.xml_to_list_text(root, './/dc:publisher')
        subject = self.xml_to_list_text(root, './/dc:subject')
        return isbn, title, creator, publisher, subject

    def fetch_book_data(self, isbn):
        _, title, creator, publisher, subject = self.fetch_book_data_NDL(isbn)
        if title=='':
            try:
                _, title, creator, publisher, subject = self.fetch_book_data_openlibrary(isbn)
            except: pass
        return isbn, title, creator, publisher, subject
    
if __name__ == "__main__":
    root = tk.Tk()
    w = root.winfo_screenwidth()    #モニター横幅取得
    h = root.winfo_screenheight()   #モニター縦幅取得
    windw_w = 700
    windw_h = 400

    root.title(u"本管理ツール")
    iconfile = 'icon.ico'
    root.iconbitmap(default=iconfile)
    root.geometry(str(windw_w)+"x"+str(windw_h)+"+"+str((w-windw_w)//2)+"+"+str((h-windw_h)//2))
    app = Application(master=root)
    app.mainloop()