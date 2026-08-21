from polars import DataFrame, read_excel, concat
from xlsxwriter import Workbook
from pathlib import Path
from re import Pattern, compile, sub, IGNORECASE
from unicodedata import normalize
from uuid import uuid4
from logging import basicConfig, getLogger, ERROR, DEBUG
from zipfile import ZIP_DEFLATED, ZipFile
from shutil import rmtree



def extract_zip_file(zip_file_path: Path, output_directory_path: Path) -> None:
    with ZipFile(file=zip_file_path, mode='r') as zip_file:
        zip_file.extractall(output_directory_path)


basicConfig(
    level=DEBUG,
    format='%(asctime)s [ %(levelname)s ]\t%(message)s',
    datefmt='%d/%m/%y %H:%M:%S %p'
)
getLogger('fastexcel').setLevel(ERROR)

logger = getLogger(__name__)

execution_path = Path(__file__).parent
base_directory_path = (execution_path / 'base')
source_directory_path = (execution_path / 'source')
data_directory_path = (execution_path / 'data')
output_directory_path = (execution_path / 'output')


for item in source_directory_path.iterdir():
    if item.suffix != '.zip':
        continue

    extract_zip_file(item, data_directory_path)



allowed_extensions = ('.xlsx', '.xlsb', '.xls', '.xlsm')
data_files = [
    (root / file)
    for root, _, files in data_directory_path.walk()
    for file in files
]

total_data_files = len(data_files)


unnamed_pattern = compile(r'^__unnamed__', IGNORECASE)


sheet_column_range = {
    1: 91,
    2: 89,
    3: 87,
    4.2: 23,
    4.3: 69,
    5: 45,
}

sheet_names = {
    1: 'OG1 - Detalle',
    2: 'OG2 - Detalle',
    3: 'OG3 - Detalle',
    4.2: 'OG4 - Detalle M',
    4.3: 'OG4 - Personal',
    5: 'OG5 - Detalle',
}

base_sheet_names = {
    1: 'og1_detail',
    2: 'og2_detail',
    3: 'og3_detail',
    4.2: 'og4_misc_detail',
    4.3: 'og4_staff',
    5: 'og5_detail',
}

base_file_path = next(base_directory_path.glob('*.xlsx'))
base_file_sheets = read_excel(source=base_file_path, sheet_id=0, read_options={'header_row': 1})

canonical_columns = {
    sheet_code: base_file_sheets[base_sheet_name].columns
    for sheet_code, base_sheet_name in base_sheet_names.items()
}


def clear_directory(directory_path: Path) -> None:
    for item in directory_path.iterdir():
        if item.is_file() or item.is_symlink():
            item.unlink()

        elif item.is_dir():
            rmtree(item)


def remove_directory(directory_path: Path) -> None:
    rmtree(directory_path)



def compress_to_zip_file() -> None:
    zip_file_path = (output_directory_path / f'{uuid4()}.zip')
    with ZipFile(file=zip_file_path, mode='w', compression=ZIP_DEFLATED, compresslevel=9) as zip_file:
        for file in data_files:
            zip_file.write(file)


def ensure_directory(directory_path: Path) -> None:
    directory_path.mkdir(parents=True, exist_ok=True)


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


def identify_sheet(name: str) -> int | float | None:
    normalized = normalize_string(name)
    for pattern, index in patterns():
        if pattern.search(normalized):
            return index
    return None


def drop_columns(dataframe: DataFrame, *, exclude: set[str] = frozenset()) -> DataFrame:
    columns_to_drop = [
        column for column in dataframe.columns
        if column in exclude or unnamed_pattern.match(column)
    ]
    return dataframe.drop(columns_to_drop)


def export_dataframe(workbook: Workbook, worksheet_name: str, dataframe: DataFrame) -> None:
    worksheet = workbook.add_worksheet(worksheet_name)
    worksheet.write_row(0, 0, dataframe.columns)

    total_rows = dataframe.height
    rows_width = len(str(total_rows))

    for row_index, row in enumerate(dataframe.iter_rows(), start=1):
        progress = f'{row_index:>{rows_width}}/{total_rows}'
        logger.debug(f'{progress:<{rows_width * 2 + 1 + 4}}{worksheet_name}')
        worksheet.write_row(row_index, 0, row)


def main():
    storage: dict[int | float, list[DataFrame]] = {}
    errors: list[dict[str, str]] = []

    for file_index, file in enumerate(data_files, 1):

        if file.suffix not in allowed_extensions:
            error = f'Archivo con extensión inválida \'{file.suffix}\''
            errors.append({'Archivo': file.name, 'Descripción': error})
            continue

        excel_file_content = read_excel(
            source=file,
            engine='calamine',
            sheet_id=0,
            raise_if_empty=False,
            read_options={'dtypes': 'string', 'header_row': 1}
        )

        for sheet_name, dataframe in excel_file_content.items():
            sheet_code = identify_sheet(sheet_name)
            if sheet_code is None:
                error = f'Hoja inválida \'{sheet_name}\''
                errors.append({'Archivo': file.name, 'Descripción': error})
                continue

            if dataframe.is_empty():
                error = 'Planilla vacía'
                errors.append({'Archivo': file.name, 'Descripción': error})
                continue


            if dataframe.shape[1] == sheet_column_range[sheet_code]:
                dataframe.columns = canonical_columns[sheet_code]
                storage.setdefault(sheet_code, []).append(dataframe)

                progress_width = len(str(total_data_files))
                progress = f'{file_index:>{progress_width}}/{total_data_files}'

                logger.debug(f'{progress:<{progress_width * 2 + 1 + 4}}{sheet_name:<28}{file.name}')
                continue

            elif dataframe.shape[1] < sheet_column_range[sheet_code]:
                error = f'{sheet_column_range[sheet_code] - dataframe.shape[1]} Columnas faltantes'

            elif dataframe.shape[1] > sheet_column_range[sheet_code]:
                error = f'{dataframe.shape[1] - sheet_column_range[sheet_code]} Columnas sobrantes'

            errors.append({'Archivo': file.name, 'Descripción': error})


    ensure_directory(output_directory_path)
    
    total_sheets = len(storage)
    groups_width = len(str(total_sheets))

    for sheet_index, (sheet_code, dataframes) in enumerate(storage.items(), 1):
        combined = concat(dataframes, how='vertical')
        combined = combined.fill_null('Sin información')
        sheet_file_path = (output_directory_path / f'{sheet_names[sheet_code]}.xlsx')

        workbook = Workbook(sheet_file_path)
        workbook.use_zip64()
        export_dataframe(workbook, sheet_names[sheet_code], combined)
        workbook.close()

        progress = f'{sheet_index:>{groups_width}}/{total_sheets}'
        logger.debug(f'{progress:<{groups_width * 2 + 1 + 4}}{sheet_names[sheet_code]:<28}{len(combined)}')

    errors_file_path = output_directory_path / 'Errores.xlsx'
    errors_workbook = Workbook(errors_file_path)
    export_dataframe(errors_workbook, 'Errores', DataFrame(errors))
    errors_workbook.close()
    
    
    
    compress_to_zip_file()

if __name__ == '__main__':
    main()