import tkinter as tk
from PIL import ImageGrab
import cv2
from tkinter import filedialog
from PIL import Image, ImageTk
import pytesseract
from docx import Document

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# creating window
root = tk.Tk()
root.title('It is simple converter')
root.geometry('650x450')


def capture():
    img = ImageGrab.grab()
    img.save('photo.jpg')
    lbl = tk.Label(root, text='Ýüklendi')
    lbl.pack(pady=20)


def capture_web():
    # Initialize the webcam
    cap = cv2.VideoCapture(0)
    # Check if the webcam is opened successfully
    if not cap.isOpened():
        print("Error: Could not open webcam")
        exit()
    # Read an image from the webcam
    ret, frame = cap.read()
    # Save the image as photo.jpg
    cv2.imwrite('photo.jpg', frame)
    # Release the webcam
    cap.release()
    # Print a message
    print("Image saved as photo.jpg")


def open_file_dialog_and_save_image():
    # Open file dialog and allow user to select an image file
    file_path = filedialog.askopenfilename(
        title="Select an Image",
        filetypes=[("Image files", "*.jpg *.jpeg *.png *.bmp *.tiff")]
    )

    if file_path:
        try:
            # Open the selected image file
            image = Image.open(file_path)

            # Save the image as photo.jpg
            image.save('photo.jpg')
            print(f"Image saved as photo.jpg")

        except Exception as e:
            print(f"An error occurred: {e}")
    else:
        print("No file selected")


def convert_to_word():
    image_path = 'photo.jpg'
    word_path = 'new.docx'
    image = Image.open(image_path)
    text = pytesseract.image_to_string(image)
    doc = Document()
    doc.add_paragraph(text)
    doc.save(word_path)


# creating button loading picture from file
load_from_file = tk.Button(root, text='Suraty ýükle', height='2', width='100', command=open_file_dialog_and_save_image
                           , font='Consolas, 12')
load_from_file.pack()

# creating button that captures an image from camera
capture_screenshoot = tk.Button(root, text='Ekrany surata al', height='2', width='100', command=capture,
                                font='Consolas, 12')
capture_screenshoot.pack()

# creating a button that takes picture from web
capture_picture = tk.Button(root, text='Web kameradan surata al', height='2', width='100', command=capture_web,
                            font='Consolas, 12')
capture_picture.pack()

convert_to_word = tk.Button(root, text='Worda geçir', height='2', width='100', command=convert_to_word,
                            font='Consolas, 12')
convert_to_word.pack()

# creating an image
image_path = 'py.jpg'
image = Image.open(image_path)

# Convert the image to a PhotoImage object
photo = ImageTk.PhotoImage(image)

# Create a Label widget to display the image
label = tk.Label(root, image=photo)
label.image = photo  # Keep a reference to avoid garbage collection
label.pack()

root.mainloop()
