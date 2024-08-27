# import os
# import xlrd
# from openpyxl import Workbook, load_workbook
# from openpyxl.utils import range_boundaries
# from openpyxl.styles import PatternFill, Border, Alignment, Font
# import uuid
# import re
#
# def convert_xls_to_xlsx(xls_file, xlsx_file):
#     # Membaca file .xls dengan xlrd
#     workbook_xls = xlrd.open_workbook(xls_file)
#     sheet_xls = workbook_xls.sheet_by_index(0)
#
#     # Membuat workbook baru untuk .xlsx dengan openpyxl
#     workbook_xlsx = Workbook()
#     sheet_xlsx = workbook_xlsx.active
#     sheet_xlsx.title = sheet_xls.name
#
#     # Menyalin semua data dari .xls ke .xlsx
#     for row in range(sheet_xls.nrows):
#         for col in range(sheet_xls.ncols):
#             cell_value = sheet_xls.cell_value(row, col)
#             sheet_xlsx.cell(row=row + 1, column=col + 1).value = cell_value
#
#     # Menyimpan file .xlsx yang baru
#     workbook_xlsx.save(xlsx_file)
#
# def unmerge_cells_but_keep_values(sheet):
#     for merged_cell_range in list(sheet.merged_cells.ranges):
#         min_col, min_row, max_col, max_row = range_boundaries(str(merged_cell_range))
#         top_left_cell_value = sheet.cell(row=min_row, column=min_col).value
#         sheet.unmerge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)
#         for row in range(min_row, max_row + 1):
#             for col in range(min_col, max_col + 1):
#                 sheet.cell(row=row, column=col).value = top_left_cell_value
#
# def remove_duplicate_words(header):
#     words = header.split()
#     seen = set()
#     result = []
#     for word in words:
#         if word not in seen:
#             seen.add(word)
#             result.append(word)
#     return " ".join(result)
#
# def merge_headers(sheet, header_rows):
#     headers = []
#     for col in range(1, sheet.max_column + 1):
#         combined_header = []
#         for row in header_rows:
#             cell_value = sheet.cell(row=row, column=col).value
#             if cell_value:
#                 combined_header.append(str(cell_value).strip())
#         combined_header_str = " ".join(combined_header)
#         unique_header = remove_duplicate_words(combined_header_str)
#         headers.append(unique_header)
#     return headers
#
# def update_header_row(sheet, header_row_num, headers):
#     for col, header in enumerate(headers, start=1):
#         cell = sheet.cell(row=header_row_num, column=col)
#         cell.value = header
#         cell.border = Border()  # Clear borders
#         cell.fill = PatternFill(fill_type=None)  # Clear background color
#         cell.alignment = Alignment(horizontal='left', vertical='top')  # Reset alignment
#         cell.font = Font()  # Reset font style
#
# def clear_extra_header_rows(sheet, header_rows):
#     for row in header_rows[:-1]:
#         for col in range(1, sheet.max_column + 1):
#             sheet.cell(row=row, column=col).value = None
#
# def auto_adjust_column_width(sheet):
#     for col in sheet.columns:
#         max_length = 0
#         column_letter = col[0].column_letter
#         for cell in col:
#             try:
#                 if cell.value:
#                     max_length = max(max_length, len(str(cell.value)))
#             except:
#                 pass
#         adjusted_width = max_length + 2
#         sheet.column_dimensions[column_letter].width = adjusted_width
#
# def remove_formulas_and_keep_values(sheet):
#     for row in sheet.iter_rows():
#         for cell in row:
#             if cell.data_type == "f":  # Memeriksa apakah sel berisi rumus
#                 try:
#                     cell.value = cell.value  # Gantikan rumus dengan nilai hasilnya
#                     cell.data_type = "n"  # Mengubah tipe data menjadi numerik atau teks sesuai hasilnya
#                 except:
#                     pass  # Jika ada masalah dengan rumus, tetap lanjut
#
# def remove_empty_rows(sheet):
#     rows_to_delete = []
#     for row in range(2, sheet.max_row + 1):  # Mulai dari baris ke-2 untuk menghindari header
#         is_empty = True
#         for col in range(1, sheet.max_column + 1):
#             cell_value = sheet.cell(row=row, column=col).value
#             if cell_value is not None and cell_value != '':
#                 is_empty = False
#                 break
#         if is_empty:
#             rows_to_delete.append(row)
#
#     for row in reversed(rows_to_delete):
#         sheet.delete_rows(row)
#
# def set_row_heights(sheet, height=30):
#     for row in sheet.iter_rows():
#         sheet.row_dimensions[row[0].row].height = height
#
# def process_file(file_path, processed_folder, image_folder):
#     file_extension = os.path.splitext(file_path)[1].lower()
#     if file_extension == '.xls':
#         xlsx_file = os.path.join(processed_folder, "converted_file.xlsx")
#         convert_xls_to_xlsx(file_path, xlsx_file)
#     else:
#         xlsx_file = file_path
#
#     workbook = load_workbook(filename=xlsx_file)
#     sheet = workbook.active
#
#     # Pertahankan nilai hasil rumus dan hapus rumus
#     remove_formulas_and_keep_values(sheet)
#
#     # Unmerge cells and keep values
#     unmerge_cells_but_keep_values(sheet)
#
#     # Remove empty rows
#     remove_empty_rows(sheet)
#
#     # Image processing
#     image_paths = []
#     image_filename_column = sheet.max_column + 1
#     sheet.cell(row=1, column=image_filename_column).value = "Image Filename"
#
#     for image in sheet._images:
#         img_filename = f"{uuid.uuid4().hex}.png"
#         img_path = os.path.join(image_folder, img_filename)
#
#         with open(img_path, "wb") as f:
#             f.write(image.ref.read())  # Menyimpan gambar ke disk
#
#         image_paths.append(img_path)
#
#         row = image.anchor._from.row
#         sheet.cell(row=row, column=image_filename_column).value = img_filename
#
#     # Handle header merging
#     header_rows = list(range(1, 4))  # Disesuaikan dengan jumlah baris header
#     headers = merge_headers(sheet, header_rows)
#     update_header_row(sheet, header_rows[0], headers)
#     clear_extra_header_rows(sheet, header_rows)
#
#     # Auto-adjust column widths
#     auto_adjust_column_width(sheet)
#
#     # Set row heights
#     set_row_heights(sheet, height=30)
#
#     output_file = os.path.join(processed_folder, "updated_" + os.path.basename(xlsx_file))
#     workbook.save(output_file)
#     return output_file, image_paths



import os
import xlwings as xw
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries
from openpyxl.styles import PatternFill, Border, Alignment, Font
import uuid
import re

def convert_xls_to_xlsx(xls_file, output_file):
    app = xw.App(visible=False)
    workbook = app.books.open(xls_file)

    sheet = workbook.sheets[0]

    # Unhide all rows and columns
    sheet.api.Rows.Hidden = False
    sheet.api.Columns.Hidden = False

    # Save the file as .xlsx
    workbook.save(output_file)
    workbook.close()
    app.quit()
    print(f"Converted '{xls_file}' to '{output_file}' with images preserved.")

def copy_with_xlwings(source_file, target_file):
    app = xw.App(visible=False)
    source_wb = None  # Inisialisasi variabel dengan None
    target_wb = None  # Inisialisasi variabel dengan None
    try:
        source_wb = app.books.open(source_file)
        target_wb = xw.Book()  # Create a new workbook

        for sheet in source_wb.sheets:
            new_sheet = target_wb.sheets.add(sheet.name)
            sheet.range("A1").expand().copy(new_sheet.range("A1"))

        target_wb.save(target_file)
    finally:
        if source_wb:
            source_wb.close()
        if target_wb:
            target_wb.close()
        app.quit()
        app.kill()  # Paksa Excel untuk menutup jika masih terbuka

def unmerge_cells_but_keep_values(sheet):
    for merged_cell_range in list(sheet.merged_cells.ranges):
        min_col, min_row, max_col, max_row = range_boundaries(str(merged_cell_range))
        top_left_cell_value = sheet.cell(row=min_row, column=min_col).value
        sheet.unmerge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                sheet.cell(row=row, column=col).value = top_left_cell_value

def remove_duplicate_words(header):
    words = header.split()
    seen = set()
    result = []
    for word in words:
        if word not in seen:
            seen.add(word)
            result.append(word)
    return " ".join(result)

def merge_headers(sheet, header_rows):
    headers = []
    for col in range(1, sheet.max_column + 1):
        combined_header = []
        for row in header_rows:
            cell_value = sheet.cell(row=row, column=col).value
            if cell_value:
                combined_header.append(cell_value.strip())
        unique_header = remove_duplicate_words(" ".join(combined_header))
        headers.append(unique_header)
    return headers

def clear_styles_for_rows(sheet, rows):
    for row in rows:
        for col in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=row, column=col)
            cell.border = Border()  # Clear borders
            cell.fill = PatternFill(fill_type=None)  # Clear background color
            cell.alignment = Alignment(horizontal='left', vertical='top')  # Reset alignment
            cell.font = Font()  # Reset font style

def update_header_row(sheet, header_row_num, headers):
    for col, header in enumerate(headers, start=1):
        cell = sheet.cell(row=header_row_num, column=col)
        cell.value = header
        cell.border = Border()  # Clear borders
        cell.fill = PatternFill(fill_type=None)  # Clear background color
        cell.alignment = Alignment(horizontal='left', vertical='top')  # Reset alignment
        cell.font = Font()  # Reset font style

def clear_extra_header_rows(sheet, header_rows):
    for row in header_rows[:-1]:
        for col in range(1, sheet.max_column + 1):
            sheet.cell(row=row, column=col).value = None

def auto_adjust_column_width(sheet, columns_with_images):
    for col in sheet.columns:
        column_letter = col[0].column_letter
        if column_letter in columns_with_images:
            continue
        max_length = 0
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[column_letter].width = adjusted_width

def normalize_columns(sheet):
    columns_to_delete = []
    for col in range(1, sheet.max_column + 1):
        for row in range(2, sheet.max_row + 1):
            cell_value = str(sheet.cell(row=row, column=col).value).strip().lower()
            if cell_value == "x":
                columns_to_delete.append(col)
                break
    for col in sorted(columns_to_delete, reverse=True):
        sheet.delete_cols(col)

def remove_formulas_and_keep_values(sheet):
    for row in sheet.iter_rows():
        for cell in row:
            if cell.data_type == "f":  # Memeriksa apakah sel berisi rumus
                cell.value = cell.value  # Gantikan rumus dengan nilai hasilnya
                cell.data_type = "n"  # Mengubah tipe data menjadi numerik atau teks sesuai hasilnya

def remove_empty_rows(sheet):
    rows_to_delete = []
    for row in range(2, sheet.max_row + 1):  # Mulai dari baris ke-2 untuk menghindari header
        is_empty = True
        for col in range(1, sheet.max_column + 1):
            cell_value = sheet.cell(row=row, column=col).value
            if cell_value is not None and cell_value != '':
                is_empty = False
                break
        if is_empty:
            rows_to_delete.append(row)

    # Hapus baris yang kosong mulai dari baris paling akhir agar index tidak berubah
    for row in reversed(rows_to_delete):
        sheet.delete_rows(row)

def remove_columns_with_values(sheet, values_to_remove):
    columns_to_delete = set()

    for col in range(1, sheet.max_column + 1):
        for row in range(1, sheet.max_row + 1):
            cell_value = str(sheet.cell(row=row, column=col).value).strip().lower()
            if cell_value in values_to_remove:
                columns_to_delete.add(col)
                break

    # Hapus kolom mulai dari kolom paling akhir agar index tidak berubah
    for col in sorted(columns_to_delete, reverse=True):
        sheet.delete_cols(col)

def remove_columns_with_images(sheet):
    columns_with_images = set()

    for image in sheet._images:
        col = image.anchor._from.col + 1
        columns_with_images.add(col)

    # Hapus kolom mulai dari kolom paling akhir agar index tidak berubah
    for col in sorted(columns_with_images, reverse=True):
        sheet.delete_cols(col)

    # Bersihkan gambar dari sheet
    sheet._images = []

def set_row_heights(sheet, height=30):
    for row in sheet.iter_rows():
        sheet.row_dimensions[row[0].row].height = height


def evaluate_formula(sheet, formula):
    # Bersihkan formula dari tanda '='
    formula = formula.lstrip('=').replace(' ', '')

    # Gantikan referensi cell dengan nilai aktual
    # Contoh, 'A1+B2' -> '10+20' jika A1=10 dan B2=20
    tokens = re.split(r'(\+|\-|\*|\/|\(|\))', formula)
    for i, token in enumerate(tokens):
        if re.match(r'^[A-Z]+[0-9]+$', token):  # Jika token adalah referensi cell
            cell_value = sheet[token].value
            if cell_value is None:
                cell_value = 0  # Perlakukan sel kosong sebagai 0
            tokens[i] = str(cell_value)

    # Gabungkan kembali formula yang sudah diubah
    evaluated_formula = ''.join(tokens)

    try:
        # Hitung hasilnya
        result = eval(evaluated_formula)
    except Exception as e:
        result = None

    return result

def copy_values_and_remove_formulas(sheet):
    for row in sheet.iter_rows():
        for cell in row:
            if cell.data_type == "f":  # Memeriksa apakah sel berisi rumus
                formula = cell.value
                formula_result = evaluate_formula(sheet, formula)  # Evaluasi rumus
                if formula_result is not None:
                    cell.value = formula_result  # Gantikan rumus dengan hasilnya
                else:
                    cell.value = "Error"  # Jika terjadi kesalahan dalam evaluasi
                cell.data_type = "n"  # Mengubah tipe data menjadi numerik atau teks sesuai hasilnya


def process_file(file_path, processed_folder, image_folder):
    file_extension = os.path.splitext(file_path)[1].lower()
    if file_extension == '.xls':
        xlsx_file = 'converted_file.xlsx'
        convert_xls_to_xlsx(file_path, xlsx_file)
    else:
        xlsx_file = 'cleaned_file.xlsx'
        copy_with_xlwings(file_path, xlsx_file)

    workbook = load_workbook(filename=xlsx_file)
    sheet = workbook.active

    # Pertahankan nilai hasil rumus dan hapus rumus
    copy_values_and_remove_formulas(sheet)

    unmerge_cells_but_keep_values(sheet)
    normalize_columns(sheet)

    # Tambahkan fitur penghapusan baris kosong
    remove_empty_rows(sheet)

    image_paths = []

    # Tentukan kolom tetap untuk menyimpan nama gambar
    image_filename_column = sheet.max_column + 1
    sheet.cell(row=1, column=image_filename_column).value = "Image Filename"

    rows_with_images = []
    columns_with_images = set()

    for image in sheet._images:
        # Generate a unique filename for each image
        img_filename = f"{uuid.uuid4().hex}.png"
        img_path = os.path.join(image_folder, img_filename)

        # Access the image's binary data
        img_stream = image.ref
        img_stream.seek(0)  # Ensure you're at the start of the file-like object

        # Save the image data to disk
        with open(img_path, "wb") as f:
            f.write(img_stream.read())  # Read from BytesIO and write to the file

        # Store the image path in the list
        image_paths.append(img_path)

        # Determine the row where the image is anchored
        row = image.anchor._from.row
        rows_with_images.append(row)
        columns_with_images.add(sheet.cell(row=row, column=image_filename_column).column_letter)

        # Write the image filename in the designated column
        sheet.cell(row=row, column=image_filename_column).value = img_filename

    if rows_with_images:
        data_start_row = min(rows_with_images)
    else:
        data_start_row = 2

    header_rows = list(range(1, data_start_row))
    headers = merge_headers(sheet, header_rows)
    header_row_num = data_start_row - 1
    update_header_row(sheet, header_row_num, headers)
    clear_styles_for_rows(sheet, header_rows)
    clear_extra_header_rows(sheet, header_rows)
    auto_adjust_column_width(sheet, columns_with_images)

    # Hapus kolom dengan nilai "1", "x", dll
    # values_to_remove = {"1", "x"}
    # remove_columns_with_values(sheet, values_to_remove)
    #
    # # Hapus kolom dengan gambar
    # remove_columns_with_images(sheet)

    # Set row heights with a specific height (e.g., 30px)
    set_row_heights(sheet, height=30)

    output_file = os.path.join(processed_folder, "updated_" + os.path.basename(xlsx_file))
    workbook.save(output_file)
    return output_file, image_paths

