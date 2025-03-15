import tkinter as tk
from tkinter import filedialog, messagebox
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import requests
import re
import threading

# Function to fetch exam data
def fetch_exam_data(exam_id):
    url = f"https://admin.genesiseud.info/api/v3/web/exam-question-answer-json/{exam_id}"
    response = requests.get(url)
    if response.status_code == 200:
        return response.json()
    else:
        raise Exception(f"Failed to fetch exam data: {response.status_code}")

# Function to generate PowerPoint
def generate_question_answer_ppt(exam):
    prs = Presentation()
    slide_layout = prs.slide_layouts[1]  # Title and Content layout
    serial_no = 1

    for question in exam["questions"]:
        slide = prs.slides.add_slide(slide_layout)
        title_content = f"{serial_no}. {clean_html_entities(question['title'])}"

        if slide.shapes.title:
            sp = slide.shapes.title
            slide.shapes._spTree.remove(sp._element)

        text_frame = slide.placeholders[1].text_frame
        text_frame.text = title_content
        text_frame.paragraphs[0].font.bold = True
        text_frame.paragraphs[0].font.size = Pt(22)
        text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

        for option in question["question_answers"]:
            p = text_frame.add_paragraph()
            p.text = clean_html_entities(option["answer"])
            p.font.size = Pt(18)
            p.alignment = PP_ALIGN.LEFT

        ans_p = text_frame.add_paragraph()
        ans_p.text = f"Ans: {question.get('answer_script', '')}"
        ans_p.font.bold = True
        ans_p.font.size = Pt(18)
        ans_p.alignment = PP_ALIGN.LEFT

        serial_no += 1

    exam_name = re.sub(r"[^a-zA-Z0-9]", "_", exam["name"] or "presentation")
    ppt_file_name = filedialog.asksaveasfilename(defaultextension=".pptx",
                                                 filetypes=[("PowerPoint Files", "*.pptx")],
                                                 initialfile=f"{exam_name}.pptx")
    if ppt_file_name:
        prs.save(ppt_file_name)
        messagebox.showinfo("Success", f"PowerPoint saved as: {ppt_file_name}")

# Utility Functions
def clean_html_entities(text):
    text = re.sub(r'&.*?;', "", text)
    text = re.sub(r'\r', "", text)
    return re.sub(r"<.*?>", "", text)

# Function to handle button click
def start_generation():
    exam_id = exam_id_entry.get().strip()
    if not exam_id:
        messagebox.showerror("Error", "Please enter an Exam ID")
        return

    def worker():
        try:
            exam_data = fetch_exam_data(exam_id)
            generate_question_answer_ppt(exam_data)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # Run in a separate thread to keep the GUI responsive
    thread = threading.Thread(target=worker)
    thread.start()

# GUI Setup
root = tk.Tk()
root.title("PowerPoint Generator")
root.geometry("400x250")

tk.Label(root, text="Enter Exam ID:").pack(pady=10)
exam_id_entry = tk.Entry(root, width=30)
exam_id_entry.pack(pady=5)

generate_button = tk.Button(root, text="Generate PowerPoint", command=start_generation)
generate_button.pack(pady=20)

root.mainloop()
