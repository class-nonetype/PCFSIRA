from unicodedata import normalize
from re import Pattern, compile, sub
from pathlib import Path
from logging import basicConfig, getLogger, ERROR, DEBUG
from zipfile import ZipFile
from shutil import rmtree
from warnings import filterwarnings
from uuid import uuid4
from pandas import ExcelFile, ExcelWriter, read_excel, DataFrame, concat


# source venv/bin/activate

basicConfig(
    level=DEBUG,
    format='%(asctime)s [ %(levelname)s ]\t%(message)s',
    datefmt='%d/%m/%y %H:%M:%S %p'
)
getLogger('fastexcel').setLevel(ERROR)
filterwarnings('ignore', message='from_arrow.*will return a Series', category=FutureWarning)

logger = getLogger(__name__)





def patterns() -> list[tuple[Pattern, (int | float)]]:
    sep = r'[\s_-]*'
    return [
        (compile(rf'^og{sep}1(?:.*)?$'), 1),                                     # og1 / og_1 / og-1 / og1_detail
        (compile(rf'^og{sep}2(?:.*)?$'), 2),                                     # og2 / og_2 / og-2 / og2_detail
        (compile(rf'^og{sep}3(?:.*)?$'), 3),                                     # og3 / og_3 / og-3 / og3_detail
        (compile(rf'^og{sep}4(?!{sep}(?:misc|staff))(?:.*)?$'), 4.1),            # og4 / og_4 / og-4 / og4_detail
        (compile(rf'^og{sep}4{sep}misc(?:.*)?$'), 4.2),                          # og4_misc_detail
        (compile(rf'^og{sep}4{sep}staff(?:.*)?$'), 4.3),                         # og4_staff
        (compile(rf'^og{sep}5(?:.*)?$'), 5),                                     # og5 / og_5 / og-5 / og5_detail
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

def read_sheet(file_path: Path, sheet_name: str, max_columns: int) -> DataFrame:
    dataframe = read_excel(io=file_path, sheet_name=sheet_name, engine='calamine', skiprows=1, dtype=str)
    return dataframe.iloc[:, :max_columns]


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



def ensure_directory(directory_path: Path) -> None:
    directory_path.mkdir(parents=True, exist_ok=True)


execution_path = Path(__file__).parent
input_directory_path = (execution_path / 'input')
data_directory_path = (execution_path / 'data')
export_directory_path = (execution_path / 'export')


ensure_directory(data_directory_path)
ensure_directory(export_directory_path)
clear_directory(data_directory_path)
clear_directory(export_directory_path)

for item in input_directory_path.iterdir():
    if item.suffix == '.zip':
        extract_zip_file(item, data_directory_path)


sheet_names = {
    1: 'OG1 - Detalle',
    2: 'OG2 - Detalle',
    3: 'OG3 - Detalle',
    4.1: 'OG4 - Detalle',
    4.2: 'OG4 - Detalle M',
    4.3: 'OG4 - Personal',
    5: 'OG5 - Detalle',
}

max_columns = {
    1: 91,
    2: 89,
    3: 87,
    4.1: 87,
    4.2: 23,
    4.3: 69,
    5: 45,
}




files = [
    (root / file)
    for root, _, files in data_directory_path.walk()
    for file in files
]
total_files = len(files)

storage: dict[int | float, list[DataFrame]] = {}
errors: list[dict[str, str]] = []



for file_index, file in enumerate(files, 1):
    with ExcelFile(path_or_buffer=file, engine='calamine') as excel_file:
        sheets = excel_file.sheet_names

    for sheet_name in sheets:
        if sheet_name.lower().startswith('i') or sheet_name.lower().startswith('d'):
            continue

        sheet_code = identify_sheet(sheet_name)
        if sheet_code is None:
            errors.append({'Archivo': file.name, 'Descripción': f'Hoja \'{sheet_name}\' inválida.'})
            continue

        dataframe = read_sheet(file, sheet_name, max_columns[sheet_code])
        storage.setdefault(sheet_code, []).append(dataframe)
        
        logger.debug('%d / %d\t%s' % (file_index, total_files, file.name))


export_batch_directory_path = export_directory_path / str(uuid4())
ensure_directory(export_batch_directory_path)


for sheet_code, dataframes in storage.items():
    combined = concat(dataframes, ignore_index=True)

    sheet_file_path = (export_batch_directory_path / f'{sheet_names[sheet_code]}.xlsx')

    with ExcelWriter(sheet_file_path, engine='xlsxwriter', engine_kwargs={'options': {'constant_memory': True}}) as excel_writer:
        combined.to_excel(excel_writer, sheet_name=sheet_names[sheet_code], index=False)

    logger.debug('%s (%d filas)' % (sheet_names[sheet_code], len(combined)))

errors_file_path = export_batch_directory_path / 'Errores.xlsx'
with ExcelWriter(errors_file_path, engine='xlsxwriter', engine_kwargs={'options': {'constant_memory': True}}) as excel_writer:
    DataFrame(errors).to_excel(excel_writer, sheet_name='Errores', index=False)
