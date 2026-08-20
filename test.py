from unicodedata import normalize
from re import Pattern, compile, sub
from concurrent.futures import ThreadPoolExecutor
from polars import DataFrame, read_excel, read_parquet, concat
from xlsxwriter import Workbook
from pathlib import Path
from dotenv import load_dotenv
from logging import basicConfig, getLogger, ERROR, DEBUG
from zipfile import ZipFile
from shutil import rmtree
from warnings import filterwarnings
from uuid import uuid4

load_dotenv()
basicConfig(
    level=DEBUG,
    format='%(asctime)s [ %(levelname)s ]\t%(message)s',
    datefmt='%d/%m/%y %H:%M:%S %p'
)
getLogger('fastexcel').setLevel(ERROR)
filterwarnings('ignore', message='from_arrow.*will return a Series', category=FutureWarning)

logger = getLogger(__name__)

execution_path = Path(__file__).parent
input_directory_path = (execution_path / 'input')
data_directory_path = (execution_path / 'data')
parts_directory_path = (execution_path / 'parts')
export_directory_path = (execution_path / 'export')

CATEGORY_NAMES: dict[int, str] = {
    22: 'og1_detail',
    23: 'og2_detail',
    24: 'og3_detail',
    25: 'og4_detail',
    26: 'og4_staff',
    27: 'og5_detail',
}


def patterns() -> list[tuple[Pattern, int]]:
    return [
        (compile(r'^og\s*1(?:\D.*)?$'), 22),                 # og1_detail
        (compile(r'^og\s*2(?:\D.*)?$'), 23),                 # og2_detail
        (compile(r'^og\s*3(?:\D.*)?$'), 24),                 # og3_detail
        (compile(r'^og\s*4(?![\s_-]*staff)(?:\D.*)?$'), 25), # og4_detail / og4_misc_detail
        (compile(r'^og\s*4[\s_-]*staff(?:\D.*)?$'), 26),     # og4_staff
        (compile(r'^og\s*5(?:\D.*)?$'), 27),                 # og5_detail
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

def deduplicate_columns(columns: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    result = []
    for c in columns:
        if c not in seen:
            seen[c] = 0
            result.append(c)
        else:
            seen[c] += 1
            result.append(f'{c}_{seen[c]}')
    return result

def ensure_directory(directory_path: Path) -> None:
    directory_path.mkdir(parents=True, exist_ok=True)


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


def process_excel_files(files: list[Path]) -> set[int]:
    """Lee cada Excel y vuelca sus hojas relevantes a fragmentos Parquet en
    disco (uno por categoria), liberando el DataFrame apenas se escribe.
    Asi la memoria usada esta acotada a los archivos en vuelo (max_workers),
    en vez de crecer con el total de archivos del ZIP."""
    total_files = len(files)
    categories_found: set[int] = set()

    def read(f: Path) -> dict[str, DataFrame]:
        return read_excel(
            source=f,
            sheet_id=0,
            raise_if_empty=False,
            read_options={'dtypes': 'string', 'header_row': 1},
        )

    with ThreadPoolExecutor(max_workers=8) as executor:
        for excel_file_index, (excel_file_path, excel_file_sheets) in enumerate(
            iterable=zip(files, executor.map(read, files)), start=1
        ):
            for key, value in excel_file_sheets.items():

                if key.lower().startswith('i') or key.lower().startswith('d'):
                    continue

                excel_file_sheet = identify_sheet(key)
                if excel_file_sheet is None:
                    logger.warning(msg=f'{key} {excel_file_path.name}')
                    continue

                logger.debug(msg=f'{excel_file_index}/{total_files} {excel_file_path.name} {excel_file_sheet}')

                value.columns = deduplicate_columns(value.columns)

                category_directory_path = parts_directory_path / str(excel_file_sheet)
                ensure_directory(category_directory_path)
                part_file_path = category_directory_path / f'{excel_file_path.stem}_{excel_file_index}.parquet'
                value.write_parquet(part_file_path)
                categories_found.add(excel_file_sheet)

            # excel_file_sheets y sus DataFrames quedan sin referencias aca
            # y se liberan antes de procesar el siguiente archivo.

    return categories_found


def export_category(category: int, export_folder_path: Path) -> None:
    category_directory_path = parts_directory_path / str(category)
    part_files = sorted(category_directory_path.glob('*.parquet'))

    dataframe = concat(
        (read_parquet(part_file) for part_file in part_files),
        how='diagonal_relaxed',
    )

    sheet_name = CATEGORY_NAMES.get(category, str(category))
    export_file_path = export_folder_path / f'{sheet_name}.xlsx'

    with Workbook(str(export_file_path)) as workbook:
        dataframe.write_excel(workbook=workbook, worksheet=sheet_name.upper())

    del dataframe


def main():
    def initialize():
        ensure_directory(data_directory_path)
        ensure_directory(export_directory_path)
        ensure_directory(parts_directory_path)

        clear_directory(data_directory_path)
        clear_directory(parts_directory_path)

        for item in input_directory_path.iterdir():
            if item.suffix == '.zip':
                extract_zip_file(item, data_directory_path)
            continue

    initialize()

    files = [
        (root / file)
        for root, _, files in data_directory_path.walk()
        for file in files
    ]

    categories = process_excel_files(files)

    export_folder_path = export_directory_path / str(uuid4())
    ensure_directory(export_folder_path)

    for category in sorted(categories):
        logger.debug(msg=f'exportando categoria {category}')
        export_category(category, export_folder_path)

    remove_directory(parts_directory_path)


if __name__ == '__main__':
    main()
