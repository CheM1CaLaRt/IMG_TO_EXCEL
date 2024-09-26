import tkinter as tk
from tkinter import filedialog, Text, messagebox
from PIL import Image, ImageTk, ImageOps
import cv2
import pytesseract
import openpyxl

# Укажите путь к исполняемому файлу Tesseract (необходимо указать для Windows)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


# Функция для загрузки и обработки изображения
def load_image():
    global img
    file_path = filedialog.askopenfilename(
        initialdir="/", title="Выберите изображение",
        filetypes=(("image files", "*.jpg *.jpeg *.png"), ("all files", "*.*"))
    )

    if file_path:
        try:
            # Загрузка и обработка изображения
            image = cv2.imread(file_path)
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)

            # Показ изображения в GUI
            img = Image.fromarray(cv2.cvtColor(thresh, cv2.COLOR_BGR2RGB))
            img = img.resize((300, 300), Image.Resampling.LANCZOS)  # Изменение размера изображения
            img_tk = ImageTk.PhotoImage(img)
            panel.config(image=img_tk)
            panel.image = img_tk

            # Распознавание текста
            recognized_text = pytesseract.image_to_string(thresh, lang='rus')
            text_box.delete(1.0, tk.END)  # Очистить текстовое поле
            text_box.insert(tk.END, recognized_text)  # Вставить распознанный текст

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось обработать изображение: {str(e)}")


# Функция для сохранения текста в Excel
def save_to_excel():
    try:
        # Получить текст из текстового поля
        recognized_text = text_box.get(1.0, tk.END).strip()

        if recognized_text:
            # Создать новый Excel-файл
            workbook = openpyxl.Workbook()
            sheet = workbook.active

            # Разделить текст по строкам
            lines = recognized_text.split('\n')

            # Разделить строки на колонки (по разделителям, если есть)
            for row_num, line in enumerate(lines, start=1):
                # Предполагается, что в таблице разделение происходит по пробелам или другим символам
                columns = line.split()  # Можно использовать другой разделитель, если это необходимо
                for col_num, value in enumerate(columns, start=1):
                    sheet.cell(row=row_num, column=col_num, value=value)

            # Сохранить в Excel-файл
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")]
            )
            if file_path:
                workbook.save(file_path)
                messagebox.showinfo("Успех", "Текст успешно сохранен в Excel!")
        else:
            messagebox.showwarning("Внимание", "Нет текста для сохранения.")

    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось сохранить в Excel: {str(e)}")


# Создание GUI
root = tk.Tk()
root.title("Оцифровка табеля в Excel")
root.geometry("600x600")

# Кнопка для загрузки изображения
btn_load = tk.Button(root, text="Загрузить изображение", padx=10, pady=5, command=load_image)
btn_load.pack()

# Панель для отображения загруженного изображения
panel = tk.Label(root)
panel.pack()

# Поле для отображения распознанного текста
text_box = Text(root, wrap='word', width=50, height=10)
text_box.pack(pady=10)

# Кнопка для сохранения текста в Excel
btn_save = tk.Button(root, text="Сохранить в Excel", padx=10, pady=5, command=save_to_excel)
btn_save.pack()

# Запуск главного цикла приложения
root.mainloop()