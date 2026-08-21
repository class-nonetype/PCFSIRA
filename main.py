import sqlite3
from unicodedata import normalize
from re import Pattern, compile, sub
from pathlib import Path
from logging import basicConfig, getLogger, ERROR, DEBUG
from zipfile import ZIP_DEFLATED, ZipFile
from shutil import rmtree
from warnings import filterwarnings
from uuid import uuid4
from pandas import ExcelFile, ExcelWriter, read_excel, DataFrame, concat, isna


# source venv/bin/activate

basicConfig(
    level=DEBUG,
    format='%(asctime)s [ %(levelname)s ]\t%(message)s',
    datefmt='%d/%m/%y %H:%M:%S %p'
)
getLogger('fastexcel').setLevel(ERROR)

logger = getLogger(__name__)

execution_path = Path(__file__).parent
input_directory_path = (execution_path / 'input')
data_directory_path = (execution_path / 'data')
export_directory_path = (execution_path / 'export')

sheet_names = {
    1:      'OG1 - Detalle',
    2:      'OG2 - Detalle',
    3:      'OG3 - Detalle',
    4.2:    'OG4 - Detalle M',
    4.3:    'OG4 - Personal',
    5:      'OG5 - Detalle',
}

usecols = {
    1:      'A:CM',
    2:      'A:CK',
    3:      'A:CI',
    4.2:    'A:W',
    4.3:    'A:BQ',
    5:      'A:AS',
}


data_files: list[Path] = [
    (root / file)
    for root, _, files in data_directory_path.walk()
    for file in files
]
total_files = len(data_files)

storage: dict[int | float, list[DataFrame]] = {}
errors: list[dict[str, str]] = []


def patterns() -> list[tuple[Pattern, (int | float)]]:
    separator = r'[\s_-]*'
    return [
        (compile(rf'^og{separator}1(?:.*)?$'), 1),                                      # og1 / og_1 / og-1 / og1_detail
        (compile(rf'^og{separator}2(?:.*)?$'), 2),                                      # og2 / og_2 / og-2 / og2_detail
        (compile(rf'^og{separator}3(?:.*)?$'), 3),                                      # og3 / og_3 / og-3 / og3_detail
        (compile(rf'^og{separator}4{separator}misc(?:.*)?$'), 4.2),                     # og4_misc_detail
        (compile(rf'^og{separator}4{separator}staff(?:.*)?$'), 4.3),                    # og4_staff
        (compile(rf'^og{separator}5(?:.*)?$'), 5),                                      # og5 / og_5 / og-5 / og5_detail
    ]

def normalize_string(s: str) -> str:
    s = normalize('NFD', s).encode('ascii', 'ignore').decode('ascii')
    return sub(r'\s+', ' ', s).strip().lower()

def identify_sheet(name: str) -> int | None:
    normalized = normalize_string(name)
    for pattern, index in patterns():
        if pattern.search(normalized):
            return index
    return None

def read_sheet(file_path: Path, sheet_name: str, usecols: str) -> DataFrame:
    return read_excel(io=file_path, sheet_name=sheet_name, engine='calamine', skiprows=1, dtype=str, usecols=usecols)

def export_dataframe(worksheet, dataframe: DataFrame, file_name: str) -> None:
    worksheet.write_row(0, 0, [str(column) for column in dataframe.columns])

    total_rows = len(dataframe)
    rows_width = len(str(total_rows))

    for row_index, row in enumerate(dataframe.itertuples(index=False, name=None), start=1):
        progress = f'{row_index:>{rows_width}}/{total_rows}'
        logger.debug(f'{progress:<{rows_width * 2 + 1 + 4}}{file_name}')
        worksheet.write_row(row_index, 0, [None if isna(value) else value for value in row])

def clear_directory(directory_path: Path) -> None:
    for item in directory_path.iterdir():
        if item.is_file() or item.is_symlink():
            item.unlink()  # Elimina el archivo
        elif item.is_dir():
            rmtree(item)

def remove_directory(directory_path: Path) -> None:
    rmtree(directory_path)

def extract_zip_file(zip_file_path: Path, output_directory_path: Path) -> None:
    with ZipFile(file=zip_file_path, mode='r') as zip_file:
        zip_file.extractall(output_directory_path)

def compress_to_zip_file() -> None:
    zip_file_path = (export_directory_path / f'{uuid4()}.zip')
    with ZipFile(file=zip_file_path, mode='w', compression=ZIP_DEFLATED, compresslevel=9) as zip_file:
        for file in data_files:
            zip_file.write(file)
    logger.debug(zip_file_path.as_posix())

def ensure_directory(directory_path: Path) -> None:
    directory_path.mkdir(parents=True, exist_ok=True)


ensure_directory(data_directory_path)
ensure_directory(export_directory_path)

clear_directory(data_directory_path)
clear_directory(export_directory_path)


for item in input_directory_path.iterdir():
    if item.suffix == '.zip':
        extract_zip_file(item, data_directory_path)








for file_index, file in enumerate(data_files, 1):
    with ExcelFile(path_or_buffer=file, engine='calamine') as excel_file:
        sheets = excel_file.sheet_names

    for sheet_name in sheets:
        if sheet_name.lower().startswith('i') or sheet_name.lower().startswith('d'):
            continue

        sheet_code = identify_sheet(sheet_name)
        if sheet_code is None:
            errors.append({'Archivo': file.name, 'Descripción': f'Hoja \'{sheet_name}\' inválida.'})
            continue

        try:
            dataframe = read_sheet(file, sheet_name, usecols[sheet_code])
        except ValueError:
            errors.append({'Archivo': file.name, 'Descripción': f'Hoja \'{sheet_name}\' con columnas faltantes.'})
            continue

        dataframe['Archivo'] = file.name

        storage.setdefault(sheet_code, []).append(dataframe)

        progress_width = len(str(total_files))
        progress = f'{file_index:>{progress_width}}/{total_files}'
        logger.debug(f'{progress:<{progress_width * 2 + 1 + 4}}{sheet_name:<28}{file.name}')


export_batch_directory_path = export_directory_path / str(uuid4())
ensure_directory(export_batch_directory_path)


# excel limita cada hoja a 1.048.576 filas (incluido el encabezado); si los
# datos exceden ese límite se reparten en hojas adicionales dentro del mismo archivo.
max_data_rows_per_sheet = 1_048_576 - 1

total_sheets = len(storage)
groups_width = len(str(total_sheets))

for sheet_index, (sheet_code, dataframes) in enumerate(storage.items(), 1):
    combined = concat(dataframes, ignore_index=True)
    combined = combined.fillna('Sin información')

    export_file_path = (export_batch_directory_path / f'{sheet_names[sheet_code]}.xlsx')
    base_sheet_name = sheet_names[sheet_code]

    with ExcelWriter(export_file_path, engine='xlsxwriter', engine_kwargs={'options': {'constant_memory': True}}) as excel_writer:
        excel_writer.book.use_zip64()
        for chunk_index, chunk_start in enumerate(range(0, max(len(combined), 1), max_data_rows_per_sheet), start=1):
            chunk = combined.iloc[chunk_start:chunk_start + max_data_rows_per_sheet]
            chunk_sheet_name = base_sheet_name if chunk_index == 1 else f'{base_sheet_name} ({chunk_index})'[:31]
            worksheet = excel_writer.book.add_worksheet(chunk_sheet_name)
            export_dataframe(worksheet, chunk, export_file_path.name)

    progress = f'{sheet_index:>{groups_width}}/{total_sheets}'
    logger.debug(f'{progress:<{groups_width * 2 + 1 + 4}}{base_sheet_name:<28}{len(combined):<14}{export_file_path.name}')

errors_file_path = export_batch_directory_path / 'Errores.xlsx'
with ExcelWriter(errors_file_path, engine='xlsxwriter') as excel_writer:
    errors_worksheet = excel_writer.book.add_worksheet('Errores')
    export_dataframe(errors_worksheet, DataFrame(errors).fillna('Sin información'), errors_file_path.name)


compress_to_zip_file()