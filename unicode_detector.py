import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from docx import Document
import os
import re

# List of special Unicode characters
special_chars = {
    0x202F: "U+202F (Narrow No-Break Space)",
    0x200B: "U+200B (Zero Width Space)",
    0x2060: "U+2060 (Word Joiner)",
    0x00A0: "U+00A0 (Non-Breaking Space)",
    0x200C: "U+200C (Zero Width Non-Joiner)",
    0x200D: "U+200D (Zero Width Joiner)",
    0x2061: "U+2061 (Function Application)"
}

# Dict to hold IntVar states for checkboxes
char_vars = {}

def analyze_file():
    # Use one directory back from current working directory
    cwd = os.getcwd()
    parent_dir = os.path.dirname(cwd)

    filepath = filedialog.askopenfilename(
        initialdir=parent_dir,
        filetypes=[("Word Documents", "*.docx")],
        title="Word-Datei auswählen"
    )
    if not filepath:
        return

    try:
        doc = Document(filepath)
        full_text = "\n".join([p.text for p in doc.paragraphs])
    except Exception as e:
        messagebox.showerror("Error", f"Could not open file:\n{e}")
        return

    output_box.delete("1.0", tk.END)  # clear old results

    # Always insert header text first
    output_box.insert(
        tk.END,
        "In der Word-Datei wurde nach Unicode Identifiern gesucht. "
        "Um schneller an die entsprechende Stelle im Dokument zu navigieren, "
        "wird der Context angezeigt. Als Context werden 20 Zeichen vor- "
        "und 20 Zeichen nach dem Unicode angegeben.\n\n"
    )

    found = False
    active_chars = {code: label for code, label in special_chars.items() if char_vars[code].get() == 1}

    for i, ch in enumerate(full_text):
        if ord(ch) in active_chars:
            # Split into sentences (handles ., !, ? followed by space)
            sentences = re.split(r'(?<=[.!?])\s+', full_text)

            # Find which sentence contains the character position
            char_count = 0
            for sentence in sentences:
                if char_count <= i < char_count + len(sentence):
                    # Mark the special character in that sentence
                    rel_pos = i - char_count
                    context = sentence[:rel_pos] + "🟥" + sentence[rel_pos:]
                    
                    output_box.insert(
                        tk.END,
                        f"Gefunden: {active_chars[ord(ch)]} an Position {i}\n"
                        f"Context: '{context}'\n\n"
                    )
                    found = True
                    break
                char_count += len(sentence) + 1  # +1 for the split delimiter

    if not found:
        output_box.insert(tk.END, "Es wurden keine ausgewählten Unicode Identifier gefunden.\n")


# GUI setup
root = tk.Tk()
root.title("Unicode Character Finder für Word-Dateien")
root.geometry("800x600")

# Frame for checkboxes
checkbox_frame = tk.LabelFrame(root, text="Zu suchende Unicode Identifier", padx=10, pady=10)
checkbox_frame.pack(fill="x", padx=10, pady=5)

# Create a checkbox for each special character
for code, label in special_chars.items():
    var = tk.IntVar(value=1)  # checked by default
    cb = tk.Checkbutton(checkbox_frame, text=label, variable=var)
    cb.pack(anchor="w")
    char_vars[code] = var

# Button to select file
btn_frame = tk.Frame(root)
btn_frame.pack(pady=10)
btn = tk.Button(btn_frame, text="Word-Datei auswählen", command=analyze_file)
btn.pack()

# Output box
output_box = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=90, height=25)
output_box.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

root.mainloop()
