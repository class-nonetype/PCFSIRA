from polars import DataFrame, read_excel, concat
from pathlib import Path
from re import Pattern, compile, sub, IGNORECASE
from unicodedata import normalize
from uuid import uuid4
from logging import basicConfig, getLogger, ERROR, DEBUG


basicConfig(
    level=DEBUG,
    format='%(asctime)s [ %(levelname)s ]\t%(message)s',
    datefmt='%d/%m/%y %H:%M:%S %p'
)
getLogger('fastexcel').setLevel(ERROR)

logger = getLogger(__name__)
execution_path = Path(__file__).parent
data_directory_path = (execution_path / 'data')
input_directory_path = (execution_path / 'input')
output_directory_path = (execution_path / 'output')



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

base_template_path = next((execution_path / 'base').glob('*.xlsx'))
base_template_sheets = read_excel(source=base_template_path, sheet_id=0, read_options={'header_row': 1})

canonical_columns = {
    sheet_code: base_template_sheets[base_sheet_name].columns
    for sheet_code, base_sheet_name in base_sheet_names.items()
}


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



def main():
    storage: dict[int | float, list[DataFrame]] = {}
    errors: list[dict[str, str]] = []

    for file_index, file in enumerate(data_files, 1):

        if file.suffix not in allowed_extensions:
            error = 'Archivo con extensión inválida'
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

    output_directory_path.mkdir(parents=True, exist_ok=True)

    total_sheets = len(storage)
    groups_width = len(str(total_sheets))

    for sheet_index, (sheet_code, dataframes) in enumerate(storage.items(), 1):
        combined = concat(dataframes, how='vertical')
        sheet_file_path = (output_directory_path / f'{sheet_names[sheet_code]}.xlsx')
        combined.write_excel(sheet_file_path, worksheet=sheet_names[sheet_code], use_zip64=True)

        progress = f'{sheet_index:>{groups_width}}/{total_sheets}'
        logger.debug(f'{progress:<{groups_width * 2 + 1 + 4}}{sheet_names[sheet_code]:<28}{len(combined)}')

    errors_file_path = output_directory_path / 'Errores.xlsx'
    DataFrame(errors).write_excel(errors_file_path, worksheet='Errores')

if __name__ == '__main__':
    main()