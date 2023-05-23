from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import os, sv_ttk
from docx import Document

# Create a window
win = Tk()
win.title("IMI-StoryWriter")
win.geometry("750x550")
document = Document()

# Colors
bg = "#1C1C1C"
accent = "#57C8FF"
surface = "#2A2A2A"
white = "white"

SIZE = 14

# Create main frame
root = ttk.Frame()
root.pack(fill='both', expand=1)

# Create a topbar
tabbar = Frame(root, bg=surface)
tabbar.pack(side='top', fill='x')

# Bottom Bar
bar = Frame(root, bg=surface)
bar.pack(side='bottom', fill='x')

home_frame = Frame(root, bg=bg)

# Home
def fun_home():
    row_export.config(bg=surface)
    row_home.config(bg=accent)

# Export
def fun_export():
    global document
    
    row_home.config(bg=surface)
    row_export.config(bg=accent)

    first_line = text.get("1.0", "1.end")
    print("First line:", first_line)
    document.add_heading(first_line, 0)

    # Get the second line
    second_line = text.get("2.0", "2.end")
    document.add_heading(f"{second_line}\n", level=4)

    # Get the remaining text
    remaining_text = text.get("3.0", "end")
    paragraphs = remaining_text.split("\n")
    for paragraph in paragraphs:
        if paragraph.strip():  # Skip empty lines
            document.add_paragraph(paragraph)

    file_path = filedialog.asksaveasfilename(defaultextension=".docx")
    document.save(file_path)
    os.startfile(file_path)

    # Clear the document
    document = Document()

# Home
Label(tabbar, text=" ", bg=surface).grid(row=0, column=0)

tab_home = Label(tabbar, text='Home', bg=surface, fg=white, font=("Product Sans", 12))
tab_home.grid(row=0, column=1)

row_home = Frame(tabbar, bg=surface, width=50, height=2)
row_home.grid(row=1, column=1, columnspan=1)

tab_home.bind("<Button-1>", lambda e: fun_home())

tab_export = Label(tabbar, text='Export', bg=surface, fg=white, font=("Product Sans", 12))
tab_export.grid(row=0, column=2)

row_export = Frame(tabbar, bg=surface, width=50, height=2)
row_export.grid(row=1, column=2, columnspan=1)

tab_export.bind("<Button-1>", lambda e: fun_export())

# Textbox
text = Text(root, bg=bg, fg=white, selectbackground=accent, selectforeground=bg, font=("Roboto Regular", SIZE), bd=0, wrap=WORD, undo=True)
text.pack(fill='both', expand=1, pady=6, padx=6)

# Word Count
def update_word_count(event=None):
    text_content = text.get("1.0", "end-1c")
    word_count = len(text_content.split())
    letter_count = len(text_content.replace(" ", ""))
    status_label.config(text=f"Word Count: {word_count}")
    letter_count_label.config(text=f"Letter Count: {letter_count}")

    # Add tag to first line
    text.tag_add("tag_first_line", "1.0", "1.end")
    text.tag_configure("tag_first_line", font=("Product Sans", 26), foreground=accent, justify='center')

    text.tag_add("tag_second_line", "2.0", "2.end")
    text.tag_configure("tag_second_line", font=("Product Sans", 18), foreground='#9FB6C6', justify='center')

text.bind("<<Modified>>", update_word_count)
text.bind("<KeyRelease>", update_word_count)

# Status Bar
status_label = Label(bar, text="Word Count: 0", bg=surface, fg=white, anchor="e", padx=10)
status_label.pack(side="left")

# Letter Count
letter_count_label = Label(bar, text="Letter Count: 0", bg=surface, fg=white, anchor="e", padx=10)
letter_count_label.pack(side="right")

# Execute
if __name__ == "__main__":
    sv_ttk.set_theme('dark')
    fun_home()
    win.mainloop()
