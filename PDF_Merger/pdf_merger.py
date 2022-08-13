import os
import tkinter as tk
from tkinter.filedialog import askopenfiles, asksaveasfile
import tkinter.messagebox as msgbox
import customtkinter as ctk
from PyPDF2 import PdfFileMerger, PdfFileReader

"""
This application provide a graphical user interface to merge pdfs using the libary PyPDF2. 
There are 2 methods for the user to choose from, depending on the circumstances they face.

Method 1 is having a folder containing all pdfs to be merged. Paste the link into application and 
provide a new file name to merge.

Method 2 is to browser for various pdfs across different directory and prompt a save-as window to select
save file location and new file name.

Room for improvement will be sorting pages and delete pdfs which were wrongly loaded in method 2. 
Also another function to be added in this application can be spliting 1 pdfs into multiple page.

Welcome all contributors for my simple application.
"""

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")


# initialize the tkinter window
class App(ctk.CTk):
    WIDTH = 600
    HEIGHT = 400

    def __init__(self):
        super().__init__()
        self.title("PDF Merger")
        self.resizable(False, False)
        self.geometry(f"{App.WIDTH}x{App.HEIGHT}")

        # define some grid spaces to be placing all widgets
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.frame_left = ctk.CTkFrame(master=self, width=100, corner_radius=3)
        self.frame_left.grid(row=0, column=0, sticky="nswe")
        self.frame_left.grid_rowconfigure(10, weight=1)
        self.frame_left.grid_columnconfigure(0, weight=1)

        self.frame_right = ctk.CTkFrame(master=self, width=100, corner_radius=3)
        self.frame_right.grid(row=0, column=1, sticky="nswe")
        self.frame_right.grid_rowconfigure(10, weight=1)
        self.frame_right.grid_columnconfigure(10, weight=1)

        # initiating some variables
        self.listofpdfs = []
        self.display_pdf = tk.StringVar()
        self.cright = u"\u00A9" "GUI Created by Jayden Ang"

        # create widgets into app
        self.lbl_pdf_path = ctk.CTkLabel(master=self.frame_right, text="Put your folder link here : ", anchor='w')
        self.lbl_new_pdf = ctk.CTkLabel(master=self.frame_right, text="New PDF Name : ", anchor='w')
        self.lbl_copyright = ctk.CTkLabel(master=self.frame_right, text=self.cright, anchor='w')
        self.lbl_lblcap = ctk.CTkLabel(master=self.frame_right, text="PDFs loaded are as below:")
        self.lbl_listofpdf = ctk.CTkLabel(master=self.frame_right, textvariable=self.display_pdf, bg_color='#f2f2f2',
                                          width=450, height=230, justify='left', wrap=450)

        self.pdf_path = ctk.CTkEntry(master=self.frame_right, width=300)
        self.new_pdf = ctk.CTkEntry(master=self.frame_right, width=300)

        self.lbl_method1 = ctk.CTkButton(master=self.frame_left, text="Merge PDFs\nin a folder",
                                         command=self.showmethod1, corner_radius=15, width=100,
                                         height=int(App.HEIGHT / 2))
        self.lbl_method2 = ctk.CTkButton(master=self.frame_left, text="Load PDFs\nfrom folders",
                                         command=self.showmethod2, corner_radius=15, width=100)
        self.export_button = ctk.CTkButton(master=self.frame_right, text="Export", command=self.merge, width=300)
        self.load_button = ctk.CTkButton(master=self.frame_right, text="Load", command=self.load_pdfs)
        self.save_button = ctk.CTkButton(master=self.frame_right, text="Save As", command=self.saveas)

        self.bind('<Escape>', lambda e: self.quit_app())

        # placing widgets into frames
        self.lbl_method1.grid(row=0, rowspan=5, sticky='nsew', pady=5, padx=5)
        self.lbl_method2.grid(row=6, rowspan=5, sticky='nsew', pady=5, padx=5)

        self.showmethod1()

    def showmethod1(self):
        # self.lbl_method1.fg_color = "#4d4dff"
        # self.lbl_method2.fg_color = "#808080"

        for widget in self.frame_right.winfo_children():
            widget.grid_forget()

        self.lbl_pdf_path.grid(row=1, column=1, pady=(150, 0), padx=10)
        self.pdf_path.grid(row=1, column=2, pady=(150, 0))
        self.lbl_new_pdf.grid(row=2, column=1, pady=(20, 0), padx=10)
        self.new_pdf.grid(row=2, column=2, pady=(20, 0))
        self.export_button.grid(row=3, column=1, columnspan=2, pady=(20, 0))

    def showmethod2(self):
        # self.lbl_method1.fg_color = "#808080"
        # self.lbl_method2.fg_color = "#4d4dff"

        for widget in self.frame_right.winfo_children():
            widget.grid_forget()

        self.lbl_lblcap.grid(row=1, column=1, pady=(20, 0), padx=10, sticky='w')
        self.lbl_listofpdf.grid(row=2, column=1, pady=(20, 0), padx=12, sticky='ew')
        self.load_button.grid(row=3, column=1, pady=(20, 0))
        self.save_button.grid(row=4, column=1, pady=(20, 0))

    def merge(self):
        try:
            merged_pdf = PdfFileMerger()
            filename = self.new_pdf.get()  # get new file name from user
            pdfpath = self.pdf_path.get()  # get folder path which consists of all pdfs to be merged
            os.chdir(pdfpath)
            for file in os.listdir(pdfpath):
                if file.endswith('.pdf') or file.endswith('.PDF'):
                    self.listofpdfs.append(file)
            for pdf in self.listofpdfs:
                merged_pdf.append(PdfFileReader(pdf, 'rb'))
            merged_pdf.write(f"{filename}.pdf")  # save the merged pdfs into 1 file under user provided name
        except FileNotFoundError:
            msgbox.showerror("Warning", "No path given!")

    def load_pdfs(self):
        # let user browse for various files across their drive directory
        path_pdf = askopenfiles(parent=self, title="Browse For PDFs..",
                                mode='r', filetypes=[("PDF", "*.pdf")])
        for item in path_pdf:
            self.listofpdfs.append(item.name)

        # display the files that been inputted into the application for merging
        for i in path_pdf:
            cur = self.display_pdf.get()
            nex = cur + '\n' + i.name
            self.display_pdf.set(nex)

    def saveas(self):
        combinedpdf = PdfFileMerger()
        for pdf in self.listofpdfs:
            combinedpdf.append(PdfFileReader(pdf, 'rb'))
        savepath = asksaveasfile(mode='w', defaultextension='.pdf', filetypes=[("PDF", "*.pdf")])
        combinedpdf.write(savepath.name)

    # exit function to bind with Esc button on keyboard
    def quit_app(self):
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()
