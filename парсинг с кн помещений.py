import os
import re
import zipfile
import fitz  # PyMuPDF
import pandas as pd

# Папка для распаковки pdf-файлов
EXTRACT_FOLDER = "./data/extracted_pdfs"
OUTPUT_EXCEL_PATH = "/Users/jojo/Desktop/парсерЕГРНдата/data/EGPN_Analysis_Output.xlsx"
os.makedirs(EXTRACT_FOLDER, exist_ok=True)

# Функция для определения типа объекта
def classify_object_type(text):
    if re.search(r'земельный участок', text, re.IGNORECASE):
        return "Земельный участок"
    elif re.search(r'здание', text, re.IGNORECASE) or re.search(r'сооружение', text, re.IGNORECASE):
        return "ОКС"
    elif re.search(r'помещение', text, re.IGNORECASE):
        return "Помещение"
    else:
        return "Неизвестный тип"

# Функция для извлечения файлов из архива
def extract_zip_files(zip_paths, extract_folder):
    pdf_files = []
    for zip_path in zip_paths:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_folder)
            pdf_files.extend(
                [os.path.join(extract_folder, file) for file in zip_ref.namelist() if file.endswith(".pdf")]
            )
    return pdf_files

# Функция для извлечения текста из PDF
def extract_text_from_pdf(pdf_path):
    text = ""
    with fitz.open(pdf_path) as pdf:
        for page in pdf:
            text += page.get_text()
    return text

# Парсинг земельного участка
def parse_land_plot(text, all_land_kn):
    land_kn = re.search(r'Кадастровый номер:\s*([\d:]+)', text).group(1) if re.search(r'Кадастровый номер:\s*([\d:]+)', text) else "н/д"
    oks_inside_match = re.search(
        r'Кадастровые\s*номера\s*расположенных\s*в\s*пределах\s*земельного\s*участка\s*объектов\s*недвижимости:\s*([\s\S]+?)(?:\n\n|\Z)',
        text, re.DOTALL | re.IGNORECASE
    )
    oks_inside_list = []
    if oks_inside_match:
        oks_inside_raw = oks_inside_match.group(1)
        oks_inside_list = [
            num.strip() for num in re.split(r'[,\s]+', oks_inside_raw) if re.match(r'^\d{2}:\d{2}:\d+:\d+$', num.strip())
        ]
        oks_inside_list = list(set(oks_inside_list) - set(all_land_kn))

    return {
        'К/н зу': land_kn,
        'S ЗУ, кв,м': re.search(r'Площадь:\s*([\d.,]+)', text).group(1).replace('.', ',') if re.search(r'Площадь:\s*([\d.,]+)', text) else "н/д",
        'Вид права на ЗУ': re.search(r'Вид,.*регистрации\s*права:\s*\d+\.\d+\s+([^\n]+)', text).group(1).strip() if re.search(r'Вид,.*регистрации\s*права:\s*\d+\.\d+\s+([^\n]+)', text) else "н/д",
        'Правообладатель ЗУ': re.search(r'Правообладатель.*\n.*\n\s*([^\n]+)', text).group(1) if re.search(r'Правообладатель.*\n.*\n\s*([^\n]+)', text) else "н/д",
        'Кадастровая стоимость ЗУ': re.search(r'Кадастровая стоимость.*?:\s*([\d\s.,]+)', text).group(1).replace(' ', '').replace('.', ',') if re.search(r'Кадастровая стоимость.*?:\s*([\d\s.,]+)', text) else "н/д",
        'Кад номера оксов внутри': oks_inside_list
    }


# Парсинг зданий и сооружений (ОКС)
def parse_oks(text, all_oks_kn):
    oks_kn = re.search(r'Кадастровый номер:\s*([\d:]+)', text).group(1) if re.search(r'Кадастровый номер:\s*([\d:]+)', text) else "н/д"
    pomesh_inside_match = re.search(
        r'Кадастровые\s*номера\s*помещений,\s*машино-мест,\s*расположенных\s*в\s*здании\s*или\s*сооружении:\s*([\s\S]+?)(?:\n\n|\Z)',
        text, re.DOTALL | re.IGNORECASE
    )
    pomesh_inside_list = []
    if pomesh_inside_match:
        pomesh_inside_raw = pomesh_inside_match.group(1)
        pomesh_inside_list = [
            num.strip() for num in re.split(r'[,\s]+', pomesh_inside_raw) if
            re.match(r'^\d{2}:\d{2}:\d+:\d+$', num.strip())
        ]
        # Убираем дубликаты
        pomesh_inside_list = list(set(pomesh_inside_list) - set(all_oks_kn))

    return {
        'К/н окс': oks_kn,
        'Вид ОКС': "Здание",
        'Назначение': re.search(r'Назначение:\s*([^\n]+)', text).group(1) if re.search(r'Назначение:\s*([^\n]+)', text) else "н/д",
        'S ОКС, кв,м': re.search(r'Площадь:\s*([\d.,]+)', text).group(1).replace('.', ',') if re.search(
            r'Площадь:\s*([\d.,]+)', text) else "н/д",
        'Вид права': re.search(
            r'Вид,\s*номер,\s*дата\s*и\s*время\s*государственной\s*регистрации\s*права:\s*\d+\.\d+\s+([^\n]+)',
            text
        ).group(1).strip() if re.search(
            r'Вид,\s*номер,\s*дата\s*и\s*время\s*государственной\s*регистрации\s*права:\s*\d+\.\d+\s+([^\n]+)',
            text
        ) else "н/д",
        'Правообладатель': re.search(r'Правообладатель.*\n.*\n\s*([^\n]+)', text).group(1) if re.search(
            r'Правообладатель.*\n.*\n\s*([^\n]+)', text) else "н/д",
        'Обременения': ", ".join(set(re.findall(r'вид:\s*([^\n]+)', text))) or "н/д",
        'Кадастровые номера помещений внутри': pomesh_inside_list  # Добавляем список кадастровых номеров помещений
    }


# Парсинг помещений
def parse_pomesh(text):
    pomesh_kn = re.search(r'Кадастровый номер:\s*([\d:]+)', text).group(1) if re.search(r'Кадастровый номер:\s*([\d:]+)', text) else "н/д"

    # Извлекаем площадь помещения
    s_pomesh = re.search(r'Площадь:\s*([\d.,]+)', text).group(1).replace('.', ',') if re.search(r'Площадь:\s*([\d.,]+)', text) else "н/д"

    # Извлекаем вид права помещения
    vid_prava_pomesh = re.search(r'Вид,\s*номер,\s*дата\s*и\s*время\s*государственной\s*регистрации\s*права:\s*\d+\.\d+\s+([^\n]+)', text).group(1) if re.search(r'Вид,\s*номер,\s*дата\s*и\s*время\s*государственной\s*регистрации\s*права:\s*\d+\.\d+\s+([^\n]+)', text) else "н/д"

    # Извлекаем правообладателя помещения
    pravoobladatel_pomesh = re.search(r'Правообладатель.*\n.*\n\s*([^\n]+)', text).group(1) if re.search(r'Правообладатель.*\n.*\n\s*([^\n]+)', text) else "н/д"

    return {
        'К/н помещения': pomesh_kn,
        'S помещения, кв,м': s_pomesh,
        'Вид права помещения': vid_prava_pomesh,
        'Правообладатель помещения': pravoobladatel_pomesh
    }
# Создание Excel-отчета с данными по ЗУ, ОКС и помещениями
def create_excel_report_with_oks(land_data, oks_data, pomesh_data, OUTPUT_EXCEL_PATH):
    rows = []

    # 1. Обрабатываем данные по земельным участкам (ЗУ)
    for land in land_data:
        # Строка с данными по ЗУ
        rows.append({
            'К/н': land['К/н зу'],
            'S ЗУ, кв,м': land['S ЗУ, кв,м'],
            'Вид права на ЗУ': land['Вид права на ЗУ'],
            'Правообладатель ЗУ': land['Правообладатель ЗУ'],
            'Кадастровая стоимость ЗУ': land['Кадастровая стоимость ЗУ'],
            'S ОКС внутри': '',
            'Кад номера оксов внутри': '',
            'ОКС внутри ЗУ': '',
            'Вид ОКС': '',
            'Назначение': '',
            'S ОКС, кв,м': '',
            'Вид права': '',
            'Правообладатель': '',
            'Обременения': '',
            'Помещение внутри ОКС': '',
            'S помещения, кв,м': '',
            'Вид права помещения': '',
            'Правообладатель помещения': ''
        })

        # Получаем список кадастровых номеров ОКС внутри ЗУ
        oks_list = land.get('Кад номера оксов внутри', [])
        oks_added = set()  # Множество для хранения добавленных кадастровых номеров ОКС внутри ЗУ
        for oks_kn in oks_list:
            if oks_kn not in oks_added:
                oks_added.add(oks_kn)
                oks_info = next((oks for oks in oks_data if oks['К/н окс'] == oks_kn), None)

                if oks_info:
                    # Добавляем строку с данными по ОКС
                    rows.append({
                        'К/н': '',
                        'S ЗУ, кв,м': '',
                        'Вид права на ЗУ': '',
                        'Правообладатель ЗУ': '',
                        'Кадастровая стоимость ЗУ': '',
                        'S ОКС внутри': '',
                        'Кад номера оксов внутри': oks_kn,
                        'ОКС внутри ЗУ': oks_info['К/н окс'],  # Заполняем только один раз
                        'Вид ОКС': oks_info.get('Вид ОКС', 'н/д'),
                        'Назначение': oks_info.get('Назначение', 'н/д'),
                        'S ОКС, кв,м': oks_info.get('S ОКС, кв,м', 'н/д'),
                        'Вид права': oks_info.get('Вид права', 'н/д'),
                        'Правообладатель': oks_info.get('Правообладатель', 'н/д'),
                        'Обременения': oks_info.get('Обременения', 'н/д'),
                        'Помещение внутри ОКС': '',
                        'S помещения, кв,м': '',
                        'Вид права помещения': '',
                        'Правообладатель помещения': ''
                    })

                    # 2. Добавляем данные по помещениям внутри ОКС
                    pomesh_kn_list = oks_info.get('Кадастровые номера помещений внутри', [])
                    for pomesh_kn in pomesh_kn_list:
                        # Находим данные о помещении
                        pomesh_info = next((p for p in pomesh_data if p['К/н помещения'] == pomesh_kn), None)

                        if pomesh_info:
                            # Добавляем строку для помещения с корректной площадью
                            rows.append({
                                'К/н': '',
                                'S ЗУ, кв,м': '',
                                'Вид права на ЗУ': '',
                                'Правообладатель ЗУ': '',
                                'Кадастровая стоимость ЗУ': '',
                                'S ОКС внутри': '',
                                'Кад номера оксов внутри': '',
                                'ОКС внутри ЗУ': oks_info['К/н окс'],
                                'Вид ОКС': '',
                                'Назначение': '',
                                'S ОКС, кв,м': '',
                                'Вид права': '',
                                'Правообладатель': '',
                                'Обременения': '',
                                'Помещение внутри ОКС': pomesh_kn,
                                'S помещения, кв,м': pomesh_info.get('S помещения, кв,м', 'н/д'),
                                'Вид права помещения': pomesh_info.get('Вид права помещения', 'н/д'),
                                'Правообладатель помещения': pomesh_info.get('Правообладатель помещения', 'н/д')
                            })
                        else:
                            # Если данные по помещению не найдены, заполняем как "н/д"
                            rows.append({
                                'К/н': '',
                                'S ЗУ, кв,м': '',
                                'Вид права на ЗУ': '',
                                'Правообладатель ЗУ': '',
                                'Кадастровая стоимость ЗУ': '',
                                'S ОКС внутри': '',
                                'Кад номера оксов внутри': '',
                                'ОКС внутри ЗУ': oks_info['К/н окс'],
                                'Вид ОКС': '',
                                'Назначение': '',
                                'S ОКС, кв,м': '',
                                'Вид права': '',
                                'Правообладатель': '',
                                'Обременения': '',
                                'Помещение внутри ОКС': pomesh_kn,
                                'S помещения, кв,м': 'н/д',
                                'Вид права помещения': 'н/д',
                                'Правообладатель помещения': 'н/д'
                            })

    # Создание DataFrame и сохранение в Excel
    df = pd.DataFrame(rows)

    try:
        df.to_excel(OUTPUT_EXCEL_PATH, index=False)
        print(f"Excel файл успешно создан: {OUTPUT_EXCEL_PATH}")
    except Exception as e:
        print(f"Ошибка при сохранении Excel-файла: {e}")



def main():
    zip_paths = [
        "./data/2099.pdf.zip",
        "./data/2139.pdf.zip",
        "./data/2140.pdf.zip",
        "./data/3417.pdf.zip",
        "./data/d8c192cd-c26e-4c16-afc3-cb798e1d2cad.zip",
        "./data/3422.pdf.zip",
        "./data/3421.pdf.zip",
        "./data/2040.pdf.zip",
        "./data/2141.pdf.zip"
    ]
    pdf_files = extract_zip_files(zip_paths, EXTRACT_FOLDER)
    all_land_kn = []
    all_oks_kn = []
    all_pomesh_kn = []
    land_data = []
    oks_data = []
    pomesh_data = []

    for pdf_file in pdf_files:
        text = extract_text_from_pdf(pdf_file)
        object_type = classify_object_type(text)

        if object_type == "Земельный участок":
            parsed_data = parse_land_plot(text, all_land_kn)
            land_data.append(parsed_data)
            all_land_kn.append(parsed_data['К/н зу'])

        elif object_type == "ОКС":
            parsed_data = parse_oks(text, all_land_kn)  # all_land_kn передаем для нахождения ОКС внутри ЗУ
            oks_data.append(parsed_data)
            all_oks_kn.append(parsed_data['К/н окс'])  # Добавляем кадастровый номер ОКС в список

        elif object_type == "Помещение":
            parsed_data = parse_pomesh(text)  # Парсим данные по помещениям
            pomesh_data.append(parsed_data)
            all_pomesh_kn.append(parsed_data['К/н помещения'])

    create_excel_report_with_oks(land_data, oks_data, pomesh_data, OUTPUT_EXCEL_PATH)


if __name__ == "__main__":
    main()
