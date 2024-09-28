from tkinter import *
from tkinter import messagebox, filedialog
import fitz as fi
import openpyxl
import pytesseract
from pdf2image import convert_from_path
from pathlib import Path
from PIL import Image
import cv2
import numpy as np

def rotate_image(image, angle):
    (h, w) = image.shape[:2]
    center = (w // 2, h // 2)
    matrix = cv2.getRotationMatrix2D(center, angle, 1.0)
    rotated_image = cv2.warpAffine(image, matrix, (w, h), flags=cv2.INTER_CUBIC)
    return rotated_image

def handwritten(images):
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    texts = []
    for image in images:
        image = np.array(image)
        image = rotate_image(image, -90)
        gray_image = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
        _, thresh_image = cv2.threshold(gray_image, 128, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        text = pytesseract.image_to_string(thresh_image, config='--psm 6')
        texts.append(text.strip())
    return texts

def iconic(pdf_path):
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    ann = []
    try:
        with fi.open(pdf_path) as doc:
            for page in range(len(doc)):
                pn = doc[page]
                pannot = pn.annots()
                if pannot:
                    for annotation in pannot:
                        if annotation.type[0] == fi.PDF_ANNOT_TEXT:
                            annot_text = annotation.info["content"]
                            ann.append(annot_text)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")
    return ann

def extract_text_from_blue_boxes(images):
    lower_blue = np.array([100, 100, 100])
    upper_blue = np.array([130, 255, 255])

    executors = []
    reviewers = []

    for image in images:
        open_cv_image = np.array(image)
        image_bgr = cv2.cvtColor(open_cv_image, cv2.COLOR_RGB2BGR)
        hsv_image = cv2.cvtColor(image_bgr, cv2.COLOR_BGR2HSV)
        mask = cv2.inRange(hsv_image, lower_blue, upper_blue)
        contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        executor_info, reviewer_info = "NA", "NA" 
        
        if contours:
            main_contour = max(contours, key=cv2.contourArea)
            x, y, w, h = cv2.boundingRect(main_contour)
            img_crop = image_bgr[y:y+h, x:x+w]
            gray_img_crop = cv2.cvtColor(img_crop, cv2.COLOR_BGR2GRAY)

            text = pytesseract.image_to_string(gray_img_crop, config='--psm 6')
            lines = text.split('\n')

            for line in lines:
                if 'By' in line:
                    if not executor_info or executor_info == "NA":
                        executor_info = line.split('By', 1)[1].strip()
                    else:
                        reviewer_info = line.split('By', 1)[1].strip()

        executors.append(executor_info)
        reviewers.append(reviewer_info)

    return executors, reviewers


def extract_red_text(images):
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    extracted_text = []
    
    lower_red1 = np.array([0, 50, 50])
    upper_red1 = np.array([10, 255, 255])
    lower_red2 = np.array([170, 50, 50])
    upper_red2 = np.array([180, 255, 255])

    for image in images:
        open_cv_image = np.array(image)
        image_bgr = cv2.cvtColor(open_cv_image, cv2.COLOR_RGB2BGR)
        hsv_image = cv2.cvtColor(image_bgr, cv2.COLOR_BGR2HSV)

        mask1 = cv2.inRange(hsv_image, lower_red1, upper_red1)
        mask2 = cv2.inRange(hsv_image, lower_red2, upper_red2)
        mask = mask1 | mask2

        red_areas = cv2.bitwise_and(image_bgr, image_bgr, mask=mask)
        gray_red_areas = cv2.cvtColor(red_areas, cv2.COLOR_BGR2GRAY)
        text = pytesseract.image_to_string(gray_red_areas, config='--psm 6')
        sentences = text.strip().split('\n')
        sentences = [sentence for sentence in sentences if sentence.strip()]
        
        extracted_text.extend(sentences)
    
    return extracted_text

def extract_designation_numbers(images):
    design_numbers = []

    for image in images:
        open_cv_image = np.array(image)
        image_bgr = cv2.cvtColor(open_cv_image, cv2.COLOR_RGB2BGR)
        h, w, _ = image_bgr.shape

        roi = image_bgr[int(h*0.7):h, int(w*0.5):w]
        gray_roi = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
        gray_roi = cv2.medianBlur(gray_roi, 3)

        _, thresh_roi = cv2.threshold(gray_roi, 150, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

        custom_config = r'--oem 3 --psm 6'
        text = pytesseract.image_to_string(thresh_roi, config=custom_config)

        designation_number = "NA"
        for word in text.split():
            if word.isdigit() and len(word) == 10:
                designation_number = word
                break

        design_numbers.append(designation_number)

    return design_numbers

import os

def process_pdfs(pdf_paths, output_excel='output.xlsx'):
    if os.path.exists(output_excel):
        workbook = openpyxl.load_workbook(output_excel)
        sheet = workbook.active
        row = sheet.max_row + 1  
    else:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Extracted Information"

        sheet["A1"] = "Sr no."
        sheet["B1"] = "Design Number"
        sheet["C1"] = "Executor"
        sheet["D1"] = "Reviewer"
        sheet["E1"] = "Comments"
        row = 2 

    for pdf_path in pdf_paths:
        images = convert_from_path(pdf_path)
        if images:
            design_numbers = extract_designation_numbers(images)
            executors, reviewers = extract_text_from_blue_boxes(images)
            red_texts = extract_red_text(images)
            handwritten_texts = handwritten(images)
            iconic_texts = iconic(pdf_path)

            comments = " ".join(red_texts + iconic_texts + handwritten_texts )

            for design_number, executor, reviewer in zip(design_numbers, executors, reviewers):
                sheet[f"A{row}"] = row-1
                sheet[f"B{row}"] = design_number
                sheet[f"C{row}"] = executor
                sheet[f"D{row}"] = reviewer
                sheet[f"E{row}"] = comments             
                row += 1
            print(design_numbers)
            print(executors)
            print(reviewers)
            print(comments)

    workbook.save(output_excel)
    print(f"Information extracted and appended to {output_excel}")

def all():
    try:
        pdf_paths = filedialog.askopenfilenames(title="Select PDFs", filetypes=[("PDF Files", "*.pdf")])
        if not pdf_paths:
            messagebox.showwarning("No files selected", "Please select one or more PDF files.")
            return
        output_excel = 'output_excel.xlsx'
        process_pdfs(pdf_paths, output_excel)
        messagebox.showinfo("Success", f"Information extracted and saved to {output_excel}")
        
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def main_screen():
    global screen
    
    screen = Tk()
    screen.geometry("400x500+600+50")
    screen.title("Pic2Sheet")
    screen.configure(background="#0a3542") 
    
    Label(text="Pic2Sheet", font=("Arial", 30), bg="#0a3542", fg="white").pack(pady=50) 

    Button(text="Process PDFs", font=("Arial", 20), height=2, width=20, bg="#5c909c", fg="white", bd=0, command=all).pack(pady=20)

    screen.mainloop()

main_screen()
