import os
import sys
import time
import pandas as pd
from PIL import Image
import pillow_heif
from urllib.parse import unquote
import shutil


def get_extensions(df, link_column):
    extensions = set()
    for link in df[link_column].dropna():
        filename = os.path.basename(unquote(link))  # имя из URL
        _, ext = os.path.splitext(filename)
        extensions.add(ext.lstrip("."))  # убираем точку
    return extensions


def handle_heic(input_path):
    heif_file = pillow_heif.read_heif(input_path)
    return Image.frombytes(heif_file.mode, heif_file.size, heif_file.data, "raw")


def handle_HEIC(input_path):
    heif_file = pillow_heif.read_heif(input_path)
    return Image.frombytes(heif_file.mode, heif_file.size, heif_file.data, "raw")


def handle_jpg(input_path):
    return Image.open(input_path)


def handle_JPG(input_path):
    return Image.open(input_path)


def handle_jpeg(input_path):
    return Image.open(input_path)


def handle_JPEG(input_path):
    return Image.open(input_path)


def handle_png(input_path):
    return Image.open(input_path)


def handle_PNG(input_path):
    return Image.open(input_path)


def convert_to_jpg(input_path, output_path):
    ext = os.path.splitext(input_path)[1]
    if ext == ".heic":
        image = handle_heic(input_path)
    elif ext == ".HEIC":
        image = handle_HEIC(input_path)
    elif ext == ".jpg":
        image = handle_jpg(input_path)
    elif ext == ".JPG":
        image = handle_JPG(input_path)
    elif ext == ".jpeg":
        image = handle_jpeg(input_path)
    elif ext == ".JPEG":
        image = handle_JPEG(input_path)
    elif ext == ".png":
        image = handle_png(input_path)
    elif ext == ".PNG":
        image = handle_PNG(input_path)
    else:
        raise ValueError(f"Неподдерживаемое расширение: {ext}")
    image = image.convert("RGB")
    image.save(output_path, "JPEG", quality=95)


def compress_to_500kb(jpg_path):
    size_limit = 1 * 1024 * 500
    quality = 95
    while os.path.getsize(jpg_path) >= size_limit and quality >= 0:
        img = Image.open(jpg_path)
        img.save(jpg_path, "JPEG", quality=quality)
        quality -= 5


# Новая функция для переименования файла
def rename(jpg_path, new_name, manual_folder):
    new_filename = new_name + ".jpg"
    new_path = os.path.join(manual_folder, new_filename)
    os.rename(jpg_path, new_path)
    return new_path


# --- Основная программа ---
# folder_name = input("Введите имя папки, где нужно обработать фото: ")
folder_name = "Фото самозапись"
df = pd.read_excel(os.path.join(folder_name, "Список.xlsx"))
link_column = "Ссылка, содержащая название файла до обработки"
name_column = "Название файла после обработки"
handled_extensions = ["jpg", "jpeg", "png", "heic"]

# Проверяем расширения
extensions = get_extensions(df, link_column)
unhandled = [ext for ext in extensions if ext not in handled_extensions]

if unhandled:
    for ext in unhandled:
        print(f"Необработанное расширение: {ext}")
    sys.exit("Найдены необработанные расширения. Скрипт остановлен.")

# Создаем папки "Непонятные фото" и "Фото после обработки" внутри folder_name, если их нет
unknown_folder = os.path.join(folder_name, "Непонятные фото")
manual_folder = os.path.join(folder_name, "Фото для ручной обработки")
output_folder = os.path.join(folder_name, "Фото после обработки")
# Удаляем старые папки, если они есть
if os.path.exists(unknown_folder):
    shutil.rmtree(unknown_folder)
if os.path.exists(manual_folder):
    shutil.rmtree(manual_folder)
if os.path.exists(output_folder):
    shutil.rmtree(output_folder)
os.makedirs(unknown_folder, exist_ok=True)
os.makedirs(manual_folder, exist_ok=True)
os.makedirs(output_folder, exist_ok=True)

# Копируем все файлы из "Фото до обработки" в "Непонятные фото" и "Фото для ручной обработки"
src_folder = os.path.join(folder_name, "Фото до обработки")
for file_name in os.listdir(src_folder):
    src_path = os.path.join(src_folder, file_name)

    dst_path_unknown = os.path.join(unknown_folder, file_name)
    dst_path_manual = os.path.join(manual_folder, file_name)

    if os.path.isfile(src_path):
        shutil.copy(src_path, dst_path_unknown)
        shutil.copy(src_path, dst_path_manual)

# Обработка файлов
for link in df[link_column]:

    # удаляем из непонятных фото
    filename = os.path.basename(unquote(link))
    path_in_unknown_folder = os.path.join(unknown_folder, filename)

    if os.path.exists(path_in_unknown_folder):
        os.remove(path_in_unknown_folder)
    else:
        print(f"Файл {filename} не найден")
        continue

    # обрабатываем в ручной папке
    try:
        path_in_manual_folder = os.path.join(manual_folder, filename)
        base_name, _ = os.path.splitext(filename)
        jpg_temp_path = os.path.join(manual_folder, base_name + ".jpg")
        convert_to_jpg(path_in_manual_folder, jpg_temp_path)
        compress_to_500kb(jpg_temp_path)

        # Получаем новое имя из столбца "Название файла после обработки"
        new_name = str(df.loc[df[link_column] == link, name_column].values[0])
        jpg_temp_path = rename(jpg_temp_path, new_name, manual_folder)
        final_path = os.path.join(output_folder, new_name + ".jpg")
        if os.path.exists(final_path):
            print(f"Файл {new_name}.jpg уже существует в {output_folder}")
            raise FileExistsError

        os.replace(jpg_temp_path, final_path)
        for f in os.listdir(manual_folder):
            if f.startswith(base_name + "."):
                os.remove(os.path.join(manual_folder, f))

    except Exception as e:
        print(
            f"Ошибка при обработке {filename}, файл лежит в 'Фото для ручной обработки': {e}"
        )

# Удаляем из "Фото для ручной обработки" те файлы, которые уже есть в "Непонятные фото"
for file_name in os.listdir(manual_folder):
    path_in_manual = os.path.join(manual_folder, file_name)
    path_in_unknown = os.path.join(unknown_folder, file_name)
    if os.path.exists(path_in_unknown):
        os.remove(path_in_manual)
