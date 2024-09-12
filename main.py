import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import sys


def resource_path(relative_path):
    """ Получает абсолютный путь к ресурсу """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def find_columns(df, search_terms):
    found_columns = {}
    for col in df.columns:
        if isinstance(col, str):
            for key, term in search_terms.items():
                if term.lower() in col.lower():
                    found_columns[key] = col
    return found_columns


def process_excel_file(file_path):
    try:
        df = pd.read_excel(file_path)

        if df.empty:
            raise ValueError("Файл Excel пустой.")

        df.columns = df.columns.str.strip()

        header_row = None
        for i, row in df.iterrows():
            if any(isinstance(item, str) and 'наименование' in item.lower() for item in row):
                header_row = i

        if header_row is None:
            raise KeyError("Заголовок таблицы 'Наименование партнера' не найден.")

        df.columns = df.iloc[header_row].str.strip()
        df = df.drop(index=range(header_row + 1)).reset_index(drop=True)

        approved_column = next((col for col in df.columns if isinstance(col, str) and 'approved' in col.lower()), None)
        peredano_column = next((col for col in df.columns if isinstance(col, str) and 'передано' in col.lower()), None)

        if not approved_column or not peredano_column:
            raise KeyError("Не найдены столбцы 'approved' или 'передано'.")

        approved_date = approved_column.split()[-1]
        peredano_date = peredano_column.split()[-1]

        search_terms = {
            'partner_name': 'наименование',
            'zalito': 'залито',
            'mapp': 'мапп',
            'check': 'chek',
            'peredano': 'передано'
        }
        columns = find_columns(df, search_terms)
        if not columns:
            raise KeyError("Не удалось найти все необходимые столбцы.")

        df = df.dropna(subset=[columns['partner_name']])

        total_peredano = df[columns['peredano']].sum()

        total_zalili = df[columns.get('zalito', 0)].sum()
        total_mapp = df[columns.get('mapp', 0)].sum()

        result = f"Отчет по крупным партнерам за {approved_date}\n"
        result += f"Поставлено в работу {int(total_peredano)} {peredano_date}\n"
        result += f"\nЗалили — {int(total_zalili)}\n"
        result += f"Маппинг — {int(total_mapp)}\n\n"

        for index, row in df.iterrows():
            partner_name = row.get(columns['partner_name'], "N/A")
            zalili = row.get(columns.get('zalito', 0), 0)
            mapping = row.get(columns.get('mapp', 0), 0)
            check = row.get(columns.get('check', 0), 0)

            result += f"{partner_name} — {int(zalili)} залили, {int(mapping)} маппинга. В статусе Check {int(check)}\n"

        return result
    except FileNotFoundError:
        messagebox.showerror("Ошибка", "Файл не найден. Пожалуйста, выберите правильный файл.")
    except ValueError as ve:
        messagebox.showerror("Ошибка", f"Ошибка при чтении файла: {str(ve)}")
    except KeyError as ke:
        messagebox.showerror("Ошибка", f"Ошибка при обработке файла: {str(ke)}")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Непредвиденная ошибка: {e}")
    return None


def show_copy_or_save_controls(file_name, content):
    for widget in root.winfo_children():
        widget.destroy()

    label_file = tk.Label(root, text=f"Файл: {file_name}", font=("Arial", 11, "bold"), bg="#F0F0F0")
    label_file.pack(pady=10)

    label = tk.Label(root, text="Отчет составлен! Выберите действие с результатом:", font=("Arial", 11), bg="#F0F0F0")
    label.pack(pady=10)

    copy_button = tk.Button(root, text="Скопировать текст", command=lambda: copy_text(content), font=("Arial", 10),
                            bg="#5DADE2", fg="white", relief="flat")
    copy_button.pack(pady=5)

    save_button = tk.Button(root, text="Сохранить в файл", command=lambda: save_file(content), font=("Arial", 10),
                            bg="#58D68D", fg="white", relief="flat")
    save_button.pack(pady=5)

    new_file_button = tk.Button(root, text="Выбрать новый файл", command=open_file, font=("Arial", 10),
                                bg="#F39C12", fg="white", relief="flat")
    new_file_button.pack(pady=5)


def copy_text(content):
    root.clipboard_clear()
    root.clipboard_append(content)
    messagebox.showinfo("Скопировано", "Текст скопирован в буфер обмена")


def save_file(content):
    save_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                             filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
    if save_path:
        with open(save_path, 'w', encoding='utf-8') as file:
            file.write(content)
        messagebox.showinfo("Успех", f"Файл успешно сохранён в: {save_path}")


def open_file():
    try:
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not file_path:
            return

        result = process_excel_file(file_path)
        if result:
            show_copy_or_save_controls(os.path.basename(file_path), result)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")


def create_app():
    global root
    root = tk.Tk()
    root.title("Alish Parser")
    root.geometry("500x250")
    root.configure(bg="#F0F0F0")

    icon_path = resource_path("icon.ico")
    root.iconbitmap(icon_path)

    root.resizable(False, False)

    title_label = tk.Label(root, text="Утилита «Alish Parser»", font=("Arial", 16, "bold"), bg="#F0F0F0")
    title_label.pack(pady=20)

    label = tk.Label(root, text="Нажмите кнопку, чтобы выбрать файл Excel для обработки.", font=("Arial", 12),
                     bg="#F0F0F0")
    label.pack(pady=20)

    button = tk.Button(root, text="Выбрать файл", command=open_file, font=("Arial", 12), bg="#5DADE2", fg="white",
                       relief="flat", padx=20, pady=5)
    button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_app()
