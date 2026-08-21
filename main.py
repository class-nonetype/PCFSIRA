from polars import DataFrame, String, col, read_excel, concat
from xlsxwriter import Workbook
from pathlib import Path
from re import Pattern, compile, sub, IGNORECASE
from unicodedata import normalize
from uuid import uuid4
from logging import basicConfig, getLogger, ERROR, DEBUG
from zipfile import ZIP_DEFLATED, ZipFile
from shutil import rmtree


def ensure_directory(directory_path: Path) -> None:
    directory_path.mkdir(parents=True, exist_ok=True)



def extract_zip_file(zip_file_path: Path, output_directory_path: Path) -> None:
    if not output_directory_path.exists(): ensure_directory(output_directory_path)

    with ZipFile(file=zip_file_path, mode='r') as zip_file:
        zip_file.extractall(output_directory_path)



def clear_directory(directory_path: Path) -> None:
    for item in directory_path.iterdir():
        if item.is_file() or item.is_symlink():
            item.unlink()

        elif item.is_dir():
            rmtree(item)



def remove_directory(directory_path: Path) -> None:
    rmtree(directory_path)



def compress_to_zip_file() -> None:
    if not output_directory_path.exists():
        ensure_directory(output_directory_path)

    zip_file_path = (output_directory_path / f'{uuid4()}.zip')
    with ZipFile(file=zip_file_path, mode='w', compression=ZIP_DEFLATED, compresslevel=9) as zip_file:
        for file in data_files:
            zip_file.write(file)



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
    4.2: 'OG4 - Detalle',
    4.3: 'OG4 - Personal',
    5: 'OG5 - Detalle',
}

export_files = {
    'OG 1 - Detalle': [1],
    'OG 2 - Detalle': [2],
    'OG 3 - Detalle': [3],
    'OG 4 - Detalle': [4.2, 4.3],
    'OG 5 - Detalle': [5],
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










def drop_columns(dataframe: DataFrame, *, exclude: set[str] = frozenset()) -> DataFrame:
    columns_to_drop = [
        column for column in dataframe.columns
        if column in exclude or unnamed_pattern.match(column)
    ]
    return dataframe.drop(columns_to_drop)


def export_dataframe(workbook: Workbook, worksheet_name: str, dataframe: DataFrame) -> None:
    worksheet = workbook.add_worksheet(worksheet_name)

    header_format = workbook.add_format({
        'bold': False,
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#4472C4',
        'font_color': '#FFFFFF',
    })
    cell_format = workbook.add_format({'border': 1})

    worksheet.set_row(0, 62)
    worksheet.write_row(0, 0, dataframe.columns, header_format)

    total_rows = dataframe.height
    rows_width = len(str(total_rows))

    for row_index, row in enumerate(dataframe.iter_rows(), start=1):
        progress = f'{row_index:>{rows_width}}/{total_rows}'
        logger.debug(f'{progress:<{rows_width * 2 + 1 + 4}}{worksheet_name}')
        worksheet.write_row(row_index, 0, row, cell_format)



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
                error = f'Hoja vacía \'{sheet_name}\''
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

    
    combined_by_code = {
        sheet_code: (
            concat(dataframes, how='vertical')
            .with_columns(col(String).str.strip_chars().replace('', None))
            .fill_null('Sin información')
        )
        for sheet_code, dataframes in storage.items()
    }

    total_files = len(export_files)
    files_width = len(str(total_files))

    ensure_directory(output_directory_path)
    clear_directory(output_directory_path)

    for file_index, (file_name, sheet_codes) in enumerate(export_files.items(), 1):
        combined_sheets = {
            sheet_code: combined_by_code[sheet_code]
            for sheet_code in sheet_codes
            if sheet_code in combined_by_code
        }
        if not combined_sheets:
            continue

        sheet_file_path = (output_directory_path / f'{file_name}.xlsx')

        workbook = Workbook(sheet_file_path)
        workbook.use_zip64()
        for sheet_code, combined in combined_sheets.items():
            export_dataframe(workbook, sheet_names[sheet_code], combined)
        workbook.close()

        progress = f'{file_index:>{files_width}}/{total_files}'
        logger.debug(f'{progress:<{files_width * 2 + 1 + 4}}{file_name:<28}{sum(len(df) for df in combined_sheets.values())}')

    errors_file_path = output_directory_path / 'Errores.xlsx'
    errors_workbook = Workbook(errors_file_path)
    export_dataframe(errors_workbook, 'Errores', DataFrame(errors))
    errors_workbook.close()
    
    
    compress_to_zip_file()

if __name__ == '__main__':
    main()