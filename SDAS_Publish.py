import os
import docx
from tkinter import *
from tkinter import filedialog, messagebox

def search_sentences_in_docx(docx_path, keywords):
    """Searches for keywords in a .docx file."""
    try:
        doc = docx.Document(docx_path)
    except Exception as e:
        print(f"Error: {docx_path} could not be read. {e}")
        return []
    results = []
    for para in doc.paragraphs:
        if all(keyword.lower() in para.text.lower() for keyword in keywords):
            if para.text.strip().endswith('.'):
                results.append((para.text.strip(), docx_path))
    return results

def search_in_directory(directory, keywords):
    """Searches for keywords in all .docx files in the specified directory."""
    all_results = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.docx'):
                docx_path = os.path.join(root, file)
                results = search_sentences_in_docx(docx_path, keywords)
                all_results.extend(results)
    return all_results

def open_directory(directory):
    """Opens the specified directory."""
    os.startfile(directory)

def show_results(event=None):
    """Displays the search results."""
    keyword = keyword_entry.get().strip()
    if len(keyword) < 5:
        messagebox.showerror("Error", "Please enter a word with at least five characters...")
        return
    keywords = [kw.strip() for kw in keyword.split(",")]
    directory = directory_entry.get().strip()
    if directory:
        results = search_in_directory(directory, keywords)
        result_text.delete(1.0, END)
        if results:
            for result, path in results:
                dirname = os.path.dirname(path)
                doc_name = os.path.basename(path)
                result_text.insert(END, f"• {result} (")
                result_text.insert(END, f"{doc_name}", (doc_name, dirname))
                result_text.insert(END, ")\n\n")
                result_text.tag_bind(doc_name, "<Button-1>", lambda e, p=dirname: open_directory(p))
                result_text.tag_config(doc_name, foreground="blue", underline=True)
        else:
            result_text.insert(END, "No results found.")
    else:
        messagebox.showerror("Error", "Please enter the directory to search in!")

def show_about():
    """Displays the About window."""
    messagebox.showinfo("About", "Strategic Document Analysis System (SDAS) version: 1.01 \n\nThis application finds sentences containing specific keywords.\n\nCoded by Kürşat Küçükay on Jan 17, 2025.")

# Create the GUI
root = Tk()
root.title("Strategic Document Analysis System - SDAS")
root.geometry("600x740")

frame = Frame(root)
frame.pack(pady=20)

keyword_label = Label(frame, text="Keywords (comma-separated):")
keyword_label.pack(pady=5)

keyword_entry = Entry(frame, width=50)
keyword_entry.pack(pady=5)
keyword_entry.bind("<Return>", show_results)

# Directory input field
directory_label = Label(frame, text="Directory to Search:")
directory_label.pack(pady=5)

directory_entry = Entry(frame, width=50)
directory_entry.pack(pady=5)
directory_entry.insert(0, r"C:\users")  # Default: "C:\users"

# Search button
search_button = Button(frame, text="Search", command=show_results)
search_button.pack(pady=10)

# Text under the search button
under_button_label = Label(frame, text="FOR EYES ONLY", font=("Arial", 10, "bold"))
under_button_label.pack()

# About button (top right corner)
about_button = Button(root, text="About", command=show_about)
about_button.pack(anchor="ne", padx=10, pady=10)

# Text box and scrollbar to show results
result_frame = Frame(root)
result_frame.pack(pady=10, padx=10, fill=BOTH, expand=True)

scrollbar = Scrollbar(result_frame)
scrollbar.pack(side=RIGHT, fill=Y)

result_text = Text(result_frame, wrap=WORD, yscrollcommand=scrollbar.set)
result_text.pack(pady=10, padx=10, fill=BOTH, expand=True)
scrollbar.config(command=result_text.yview)

# Text at the bottom right corner
italic_label = Label(root, text="finds it, brings it...", font=("Arial", 10, "italic"), anchor="se")
italic_label.pack(side=BOTTOM, anchor="se", pady=5, padx=5)

root.mainloop()
