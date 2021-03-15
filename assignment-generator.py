from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename

if __name__ == "__main__":
    Tk().withdraw()
    template = askopenfilename(title="Select Word template file...", filetypes=[("Docx files", "*.docx")])
    if template == "":
        exit(0)
    datafile = askopenfilename(title="Select Excel data file...", filetypes=[("Excel files", "*.xlsx")])
    if datafile == "":
        exit(0)
    savefile = asksaveasfilename(title="Select save directory...", filetypes=[("Docx files", "*.docx"), ("All Files", "*.*")],
                                 defaultextension='.docx', initialfile='output.docx')
    if savefile == "":
        exit(0)

    print(template, datafile, savefile)
