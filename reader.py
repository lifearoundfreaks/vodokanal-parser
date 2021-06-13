from dataclasses import dataclass
from collections import defaultdict


from const import (
    SHARED_DEPARTMENTS,
    REGIONAL_DEPARTMENTS,
    JOB_SECTIONS,
    SECTION_NUMBERS,
    CONTRACT_START_TRIGGER,
    DATA_START_ROW,
    SECTION_TEMPLATE,
    TYPE_TEMPLATE,
    COST_COLUMN,
    CONTRACT_NUMBER_SIGNIFIER,
)


@dataclass
class Contract():
    job_type: str
    job_section: str
    department: str
    number: str
    address: str
    job_name: str


class Reader:

    MAIN_CONTENT_COLUMN = 0
    EMPTY_COL_SIGNIFIER = 'nan'
    ADDRESS_COLUMN_ID = 2
    INITIAL_ROW = 0

    def __init__(self, row_iter, month):

        self.iter = row_iter
        self.month = month
        self.current_row = self.INITIAL_ROW

    def get_next(self):

        self.current_row, row = next(self.iter)

        return row

    def strip_spaces(self, text):

        return " ".join(filter(None, text.split()))

    def get_main_content(self, row):

        return self.strip_spaces(str(row[self.MAIN_CONTENT_COLUMN]))

    def get_address(self, row):

        return self.strip_spaces(str(row[self.ADDRESS_COLUMN_ID]))

    def parse_contract_data(self):

        self.get_next()

        department_info = self.get_main_content(self.get_next())
        split_department = department_info.split()
        job_type = " ".join(split_department[:-2])
        job_section = JOB_SECTIONS[job_type]
        job_type = job_type.capitalize()
        department = " ".join(split_department[-2:])

        shared_department = self.get_main_content(self.get_next())
        department = SHARED_DEPARTMENTS[shared_department] \
            if shared_department != self.EMPTY_COL_SIGNIFIER \
            else REGIONAL_DEPARTMENTS[department]

        number = self.get_main_content(
            self.get_next()
        ).split(CONTRACT_NUMBER_SIGNIFIER)[-1]

        address = self.get_address(self.get_next())

        job_name = []
        while True:

            string = self.get_main_content(self.get_next())

            if self.month in string:

                return Contract(
                    job_type, job_section, department, number, address,
                    ", ".join(job_name)
                )

            job_name.append(string)

    def get_contracts(self):

        contracts = []
        add_contract = contracts.append

        try:
            while True:

                if self.get_main_content(
                    self.get_next()
                ) == CONTRACT_START_TRIGGER:
                    add_contract(self.parse_contract_data())

        except StopIteration:

            return contracts


class ContractSorter:

    def __init__(self, contracts):

        self.sections = {
            section: defaultdict(list) for section in JOB_SECTIONS.values()
        }
        self.types = {value: key for key, value in JOB_SECTIONS.items()}

        for contract in contracts:
            self.sections[contract.job_section][contract.department].append(
                contract
            )
    
    def get_dataframe_data(self):

        result = []
        add = result.append

        current_row = DATA_START_ROW

        for section, departments in sorted(
            self.sections.items(), key=lambda item: SECTION_NUMBERS[item[0]]
        ):
            section_number = SECTION_NUMBERS[section]

            add(['', '', SECTION_TEMPLATE.format(
                name=section, number=section_number
            )])

            add(['', '', TYPE_TEMPLATE.format(
                name=JOB_SECTIONS[self.types[section]], number=section_number
            )])

            current_row += 2

            for department, contracts in departments.items():

                start = current_row
                end = current_row = current_row + len(contracts)

                sum_str = f'&SUM({COST_COLUMN}{start}:{COST_COLUMN}{end})'

                add(['', '', f'="{department}: "{sum_str}'])

                for contract in contracts:

                    add([
                        '',
                        contract.number,
                        contract.address,
                        contract.job_name,
                        '',
                        '',
                        '',
                        '',
                        '',
                        '',
                        '',
                        '',
                        contract.department,
                    ])

                add([])
                current_row += 2

            add([])
            current_row += 1

        return result
