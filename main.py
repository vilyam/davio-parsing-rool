import contextlib
import csv
import datetime
import os
import re
import sys
from abc import abstractmethod, ABC
import logging
from logging import handlers, Logger
from typing import Tuple

from docx import Document
from docx.table import _Row

from transliterate import get_translit_function
from transliterate.base import TranslitLanguagePack, registry

__author__ = 'Ivan Viliamov'
__copyright__ = '2021 Ivan Viliamov'
__license__ = 'GPL 2.0/LGPL 2.1'

start_date_str = datetime.datetime.now().strftime("%y%m%d_%H%M%S")

ch = logging.StreamHandler(sys.stdout)
log_format = logging.Formatter("%(asctime)s - %(levelname)s - %(name)s: %(message)s")
ch.setFormatter(log_format)


class UkrainianDeutschLanguagePack(TranslitLanguagePack):
    __title__ = 'transliterate.contrib.languages.uk_de.translit_language_pack'

    """Language pack for Ukrainian to Deutsch language.

    See `https://de.wikipedia.org/wiki/Ukrainische_Sprache#Alphabet` for details.
    """
    language_code = "uk_de"
    language_name = "Ukrainian Deutsch"
    character_ranges = ((0x0400, 0x04FF), (0x0500, 0x052F))
    mapping = (
        u"abwhgdesyijklmnoprstufz'ABWHGDESYIJKLMNOPRSTUFZ'",
        u"абвгґдезиійклмнопрстуфцьАБВГҐДЕЗИІЙКЛМНОПРСТУФЦЬ",
    )
    pre_processor_mapping = {
        u"je": u"є",
        u"ji": u"ї",
        u"ch": u"х",
        u"tsch": u"ч",
        u"schtsch": u"щ",
        u"ju": u"ю",
        u"ja": u"я",
        u"Je": u"Є",
        u"Ji": u"Ї",
        u"Ch": u"Х",
        u"Tsch": u"Ч",
        u"Schtsch": u"Щ",
        u"Ju": u"Ю",
        u"Ja": u"Я"
    }
    reversed_specific_pre_processor_mapping = {
        u"ш": u"sch",
        u"Ш": u"Sch",
        u"ж": u"sch",
        u"Ж": u"Sch",
        u"з": u"s",
        u"З": u"S",
        u"ь": u"",
        u"Ь": u"",
    }


registry.register(UkrainianDeutschLanguagePack)


def get_logger(logger_name) -> Logger:
    lg = logging.getLogger(logger_name)
    lg.setLevel(logging.DEBUG)

    if not os.path.exists('log'):
        os.makedirs('log')

    logfile = f"log/{logger_name}_{start_date_str}.log"

    if logger_name:
        lf = logging.Formatter("%(message)s")

        fh = handlers.RotatingFileHandler(logfile)
        fh.setFormatter(lf)

        lg.addHandler(fh)
    else:
        fh = handlers.RotatingFileHandler(logfile)
        fh.setFormatter(log_format)

        lg.addHandler(fh)
        lg.addHandler(ch)

    return lg


log = get_logger('')


class SubstitutionDictionary:
    def __init__(self, file_name, path='.'):

        self.file_name = file_name
        self.path = path

        self.dictionary = {}
        self.replaced_set = set()

        self.fill_substitutions()

    def fill_substitutions(self):
        with os.scandir(self.path) as substitution_dir:
            for f in substitution_dir:
                if f.name == self.file_name and f.is_file():
                    with open(f.path, 'r') as file:
                        data = file.readlines()
                        for line in data:
                            pair = line.strip().split()
                            if len(pair) == 1:
                                pair.append(pair[0])
                            self.dictionary[pair[0]] = pair[1]

        log.info(f"There are {len(self.dictionary)} substitutions for {self.file_name}")

    def replace(self, string: str) -> Tuple[bool, str]:
        if string in self.dictionary:
            new_string = self.dictionary[string]
            self.replaced_set.add(new_string)

            return True, new_string
        else:
            return False, string


class Declension(ABC):
    def __init__(self, logger_name):
        self.log = get_logger(logger_name)
        self.consonants_rule_u = ['б', 'в', 'д', 'з', 'н', 'п', 'р', 'с', 'т', 'ф', 'ч', 'х', 'ш', 'ц']

        self.names_set_filtered = set()
        self.names_set_replaced = set()
        self.names_set_not_filtered = set()

        self.replaced_dictionary_cache = {}

    @property
    @abstractmethod
    def substitution_dictionary(self) -> SubstitutionDictionary:
        pass

    @abstractmethod
    def replace_suffix(self, in_str: str) -> Tuple[bool, str]:
        pass

    @abstractmethod
    def check_exclusion_rules(self, in_str: str) -> bool:
        pass

    @abstractmethod
    def check_to_filter_after_all(self, in_str: str) -> bool:
        pass

    @staticmethod
    def replace(string, removal, replacement):
        reverse_removal = removal[::-1]
        reverse_replacement = replacement[::-1]

        return string[::-1].replace(reverse_removal, reverse_replacement, 1)[::-1]

    def from_genitive_to_nominative_case(self, in_str: str) -> str:
        # 1. if already cached - no need to apply transformation
        if in_str in self.replaced_dictionary_cache:
            new_str = self.replaced_dictionary_cache[in_str]

            # log.debug(f"[{in_str}] replaced to [{new_str}] # cached")
        else:
            replaced, new_str = self.substitution_dictionary.replace(in_str)

            # 2. dictionary predefined exclusion
            if replaced:
                self.replaced_dictionary_cache[in_str] = new_str

                self.log.debug(f"[{in_str}] replaced to [{new_str}] # by dict")
            else:
                # 3.1 check for rule exclusions
                if self.check_exclusion_rules(in_str):
                    self.names_set_filtered.add(in_str)
                    self.replaced_dictionary_cache[in_str] = in_str

                    self.log.debug(f"[{in_str}] is exclusion")

                # 3.2 apply replacement based on algorithm
                else:
                    replaced, new_str = self.replace_suffix(in_str)
                    if replaced:
                        self.names_set_replaced.add(new_str)
                        self.replaced_dictionary_cache[in_str] = new_str

                        self.log.debug(f"[{in_str}] replaced to [{new_str}]")

                    else:
                        # 3.3 filter after all replacements
                        if self.check_to_filter_after_all(in_str):
                            self.names_set_filtered.add(in_str)

                            self.log.debug(f"[{in_str}] filtered by rules")
                        else:
                            self.names_set_not_filtered.add(in_str)

                            self.log.debug(f"[{in_str}] not find any replace rule. going to Not Filtered")

        return new_str


class NameDeclension(Declension):
    def __init__(self):
        super().__init__('name')

        self.substitution_dic = SubstitutionDictionary("names.txt", path="./exc")

    @property
    def substitution_dictionary(self) -> SubstitutionDictionary:
        return self.substitution_dic

    def check_to_filter_after_all(self, in_str: str) -> bool:
        return False

    def replace_suffix(self, in_str: str) -> Tuple[bool, str]:
        new_s = in_str
        is_matched = False
        suffix_1 = in_str[-1:]
        suffix_2 = in_str[-2:]

        if suffix_1 == "у":
            let = in_str[-2]
            if let in self.consonants_rule_u or let in ['м', 'н', 'л', 'к']:
                new_s = self.replace(in_str, let + "у", let)
                is_matched = True

        elif suffix_1 == "і":
            let = in_str[-2]
            if let in ['д', 'м', 'н', 'ш', 'т', 'р']:
                new_s = self.replace(in_str, let + "і", let + "а")
                is_matched = True
            elif let == "ц":
                new_s = self.replace(in_str, "ці", "ка")
                is_matched = True

        elif suffix_1 == "ю":
            if suffix_2 == "рю":
                new_s = self.replace(in_str, "рю", "рь")
                is_matched = True
            elif suffix_2 == "ею":
                new_s = self.replace(in_str, "ею", "ей")
                is_matched = True
            elif suffix_2 == "ію":
                new_s = self.replace(in_str, "ію", "ій")
                is_matched = True
            elif suffix_2 == "лю":
                new_s = self.replace(in_str, "лю", "ль")
                is_matched = True

        elif suffix_1 == "ї":
            if suffix_2 == "ії":
                new_s = self.replace(in_str, "ії", "ія")
                is_matched = True
            elif suffix_2 == "еї":
                new_s = self.replace(in_str, "еї", "ея")
                is_matched = True

        return is_matched, new_s

    def check_exclusion_rules(self, in_str: str) -> bool:
        suffix_1 = in_str[-1]

        if suffix_1 in ['я', 'й', 'а', 'н', 'п', 'о', 'с', 'ь', '.']:
            return True
        elif len(in_str) == 1:
            return True
        else:
            return False


class SurnameDeclension(Declension):
    def __init__(self):
        super().__init__('surname')

        self.consonants = ['б', 'в', 'г', 'ґ', 'д', 'ж', 'з', 'к', 'л', 'м', 'н', 'п', 'р', 'с', 'т', 'ф', 'х', 'ц',
                           'ч', 'ш',
                           'щ', 'ь']

        self.substitution_dic = SubstitutionDictionary("surnames.txt", path="./exc")

    @property
    def substitution_dictionary(self) -> SubstitutionDictionary:
        return self.substitution_dic

    def replace_suffix(self, in_str: str) -> Tuple[bool, str]:
        new_s = in_str
        is_matched = False
        suffix_1 = in_str[-1:]
        suffix_2 = in_str[-2:]
        suffix_3 = in_str[-3:]

        if suffix_2 == "ій":
            new_s = self.replace(in_str, "ій", "а")
            is_matched = True

        elif suffix_3 == "ому":
            let = in_str[-4]
            if let == 'ь':
                new_s = self.replace(in_str, "ьому", "ій")
            else:
                new_s = self.replace(in_str, "ому", "ий")
            is_matched = True

        elif suffix_2 == "ку":
            let = in_str[-3]
            if let in self.consonants or let == "й":
                new_s = self.replace(in_str, "ку", "ко")
            else:
                new_s = self.replace(in_str, "ку", "к")
            is_matched = True

        elif suffix_1 == "у":
            let = in_str[-2]
            if let in self.consonants_rule_u:
                new_s = self.replace(in_str, let + "у", let)
                is_matched = True

        elif suffix_1 == "ю":
            if suffix_2 == "ію":
                new_s = self.replace(in_str, "ію", "ій")
                is_matched = True
            elif suffix_2 == "аю":
                new_s = self.replace(in_str, "аю", "ай")
                is_matched = True
            elif suffix_2 == "ню":
                new_s = self.replace(in_str, "ню", "нь")
                is_matched = True
            elif suffix_2 == "лю":
                new_s = self.replace(in_str, "лю", "ль")
                is_matched = True
            elif suffix_2 == "дю":
                new_s = self.replace(in_str, "дю", "дь")
                is_matched = True
            elif suffix_2 == "рю":
                new_s = self.replace(in_str, "рю", "р")
                is_matched = True
            elif suffix_2 == "цю":
                let = in_str[-3]
                if let == 'е':
                    new_s = self.replace(in_str, "цю", "ць")
                elif let == 'й':
                    new_s = self.replace(in_str, "йцю", "єць")
                else:
                    new_s = self.replace(in_str, "цю", "ець")
                is_matched = True

        elif suffix_1 == "і":
            if suffix_2 == "бі":
                new_s = self.replace(in_str, "бі", "ба")
                is_matched = True
            elif suffix_2 == "ді":
                new_s = self.replace(in_str, "ді", "да")
                is_matched = True

        return is_matched, new_s

    def check_exclusion_rules(self, in_str: str):
        suffix_1 = in_str[-1]
        suffix_2 = in_str[-2:]

        if suffix_1 in self.consonants or suffix_1 in ['о'] or suffix_2 in ['ко']:
            return True
        else:
            return False

    def check_to_filter_after_all(self, in_str: str) -> bool:
        suffix_1 = in_str[-1]
        if suffix_1 == 'й':
            return True
        else:
            return False


class CSVWriter:
    def __init__(self):
        if not os.path.exists('out'):
            os.makedirs('out')

        self.file_csv_ok = open('out/out.csv', 'w', newline='')
        self.file_csv_warnings = open('out/out_warnings.csv', 'w', newline='')

        fieldnames_ok = ['filename', 'case_number', 'document_number', 'date', 'sender_surname', 'sender_surname_de',
                         'sender_name', 'sender_name_de', 'receiver_surname', 'receiver_surname_de', 'receiver_name',
                         'receiver_name_de', 'locality', 'locality_de', 'township', 'township_de', 'sheet_number',
                         'notice']
        self.csv_ok = csv.DictWriter(self.file_csv_ok, fieldnames=fieldnames_ok)

        fieldnames_warnings = ['filename', 'case_number', 'document_number', 'date', 'addressees', 'addressees_de',
                               'locality', 'locality_de', 'township', 'township_de', 'sheet_number', 'notice']
        self.csv_warnings = csv.DictWriter(self.file_csv_warnings, fieldnames=fieldnames_warnings)

        self.csv_ok.writeheader()
        self.csv_warnings.writeheader()

    def ok(self, row):
        self.csv_ok.writerow(row)

    def warn(self, row):
        self.csv_warnings.writerow(row)

    def close(self):
        if self.file_csv_ok:
            self.file_csv_ok.close()
            self.file_csv_ok = None

        if self.file_csv_warnings:
            self.file_csv_warnings.close()
            self.file_csv_warnings = None


class LetterDocsParser:
    def __init__(self, docs_path):
        self.surname_dec = SurnameDeclension()
        self.name_dec = NameDeclension()

        self.docs_path = docs_path

        self.translit_uk_de = get_translit_function('uk_de')

    def is_interm_header(self, r: _Row):
        not_empty_cell_count = 0
        for c in r.cells:
            if c.text:
                not_empty_cell_count += 1
            if not_empty_cell_count > 1:
                return False
        return True

    def is_locality(self, r: _Row):
        for c in r.cells:
            if c.text.lower().find("район") > 0:
                return True
        return False

    def extract_interm_header(self, r: _Row):
        for c in r.cells:
            if c.text:
                return c.text

    def is_junk_row(self, r: _Row):
        txt_1 = r.cells[0].text.strip()
        txt_2 = r.cells[1].text.strip()
        txt_5 = r.cells[4].text.strip()

        if txt_1 == "1" and txt_2 == "2" and txt_5 == "5":
            return True

        if txt_2 == "Дата документа" and txt_5 == "Примітка":
            return True

    def translit_uk_to_de(self, str_cyr):
        return self.translit_uk_de(str_cyr, reversed=True)

    def normalize_text(self, s: str):
        normalized = s.strip()
        normalized = normalized.replace("\n", " ")
        normalized = ' '.join(normalized.split())
        return normalized

    def parse_and_transform(self):
        with contextlib.closing(CSVWriter()) as csv_writer:
            with os.scandir(self.docs_path) as it:
                for entry in it:
                    if entry.name.endswith(".docx") and '~' not in entry.name and entry.is_file():
                        file_name = entry.name
                        log.info(f"reading file {file_name}")

                        case_number = ''
                        digits = re.findall(r'\d+', file_name)
                        if digits:
                            case_number = digits[0]

                        word_doc = Document(entry.path)

                        for table in word_doc.tables:
                            locality_cyr = township_cyr = ""

                            for row in table.rows:
                                if self.is_junk_row(row):
                                    continue

                                if self.is_interm_header(row):
                                    if self.is_locality(row):
                                        locality_cyr = self.normalize_text(self.extract_interm_header(row))
                                        continue
                                    else:
                                        township_cyr = self.normalize_text(self.extract_interm_header(row))
                                        if township_cyr.lower().find("вінниця") > 0:
                                            locality_cyr = "-"

                                    continue

                                document_number = self.normalize_text(row.cells[0].text)
                                date = self.normalize_text(row.cells[1].text)
                                sheet_number = self.normalize_text(row.cells[3].text)
                                note_cyr = self.normalize_text(row.cells[4].text)

                                obj_base = {
                                    'filename': file_name,
                                    'case_number': case_number,  # номер справи
                                    'document_number': document_number,  # номер документу (колонка № з/п)
                                    'date': date,  # дата,
                                    'locality': locality_cyr,  # район українською
                                    'locality_de': self.translit_uk_to_de(locality_cyr),  # район транслітерація
                                    'township': township_cyr,  # село українською
                                    'township_de': self.translit_uk_to_de(township_cyr),  # село трансліт
                                    'sheet_number': sheet_number,  # «№№ аркушів»
                                    'notice': note_cyr  # Примітка
                                }

                                addressees_cyr = row.cells[2].text
                                addressees_cyr = addressees_cyr.replace("-", "")
                                addressees_cyr = self.normalize_text(addressees_cyr)

                                addressees_list = addressees_cyr.split()
                                if len(addressees_list) != 4:
                                    log.warning(f"{addressees_list}")

                                    warn_obj_part = {
                                        'addressees': addressees_cyr,
                                        'addressees_de': self.translit_uk_to_de(addressees_cyr),
                                    }

                                    warn_obj = {**obj_base, **warn_obj_part}

                                    csv_writer.warn(warn_obj)

                                    obj_part = {
                                        'sender_surname': '',
                                        'sender_surname_de': '',
                                        'sender_name': '',
                                        'sender_name_de': '',
                                        'receiver_surname': '',
                                        'receiver_surname_de': '',
                                        'receiver_name': '',
                                        'receiver_name_de': ''
                                    }

                                    obj = {**obj_base, **obj_part}

                                    csv_writer.ok(obj)

                                    continue

                                sender_surname = addressees_list[0]
                                sender_name = addressees_list[1]

                                receiver_surname = addressees_list[2]
                                receiver_surname_r = self.surname_dec.from_genitive_to_nominative_case(receiver_surname)

                                receiver_name = addressees_list[3]
                                receiver_name_r = self.name_dec.from_genitive_to_nominative_case(receiver_name)

                                obj_part = {
                                    'sender_surname': sender_surname,  # прізвище відправника укр
                                    'sender_surname_de': self.translit_uk_to_de(sender_surname),  # прізвище відправника транслітерація
                                    'sender_name': sender_name,  # ім’я відправника укр
                                    'sender_name_de': self.translit_uk_to_de(sender_name),  # ім’я відправника транслітерація
                                    'receiver_surname': receiver_surname_r,  # прізвище отримувача укр
                                    'receiver_surname_de': self.translit_uk_to_de(receiver_surname_r),  # прізвище отримувача транслітерація
                                    'receiver_name': receiver_name_r,  # ім’я отримувача укр
                                    'receiver_name_de': self.translit_uk_to_de(receiver_name_r),  # ім’я отримувача транслітерація
                                }

                                obj = {**obj_base, **obj_part}

                                csv_writer.ok(obj)

    def print_stat(self):
        print("----------------")
        log.info(f"Replaced Surnames: {len(self.surname_dec.names_set_replaced)}\n{self.surname_dec.names_set_replaced}")

        log.info(f"Replaced Names: {len(self.name_dec.names_set_replaced)}\n{self.name_dec.names_set_replaced}")

        log.info(f"Filtered Surnames: {len(self.surname_dec.names_set_filtered)}\n{self.surname_dec.names_set_filtered}")

        log.info(f"Filtered Names: {len(self.name_dec.names_set_filtered)}\n{self.name_dec.names_set_filtered}")

        log.info(f"Not Filtered Surnames: {len(self.surname_dec.names_set_not_filtered)}\n{self.surname_dec.names_set_not_filtered}")

        log.info(f"Not Filtered Names: {len(self.name_dec.names_set_not_filtered)}\n{self.name_dec.names_set_not_filtered}")
        print("----------------")


parser = LetterDocsParser('./in')
parser.parse_and_transform()
parser.print_stat()
