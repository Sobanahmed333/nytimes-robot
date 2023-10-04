import re
import datetime
# import dateutil.parser
import time

from RPA.Excel.Files import Files
from RPA.HTTP import HTTP
from RPA.Browser.Selenium import Selenium
from selenium.common import NoSuchElementException
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By

from elements import *


class NYTimesNewsRobot:
    __browser = Selenium()
    __http = HTTP()
    __files = Files()

    def __init__(self, phrase='pakistan'):
        self.__phrase = phrase

    def __open_website(self):
        self.__browser.auto_close = False
        self.__browser.open_available_browser(URL)
        self.__browser.maximize_browser_window()

    def __close_modal(self):
        try:
            self.__browser.wait_until_element_is_visible('''//span[text()="We've updated our terms"]''')
            self.__browser.click_button(CONTINUE_BUTTON)
        except Exception as e:
            pass

    def __search_phrase(self):
        self.__browser.click_button_when_visible(SEARCH_BUTTON)
        self.__browser.input_text_when_element_is_visible(SEARCH_INPUT, self.__phrase)
        self.__browser.click_button(SEARCH_SUBMIT)

    def __apply_date_filter(self, n_prev_months):
        date_div = self.__browser.find_element(DATE_DIV)
        date_div.click()
        specific_dates_elem = date_div.find_element(by=By.XPATH, value=SPECIFIC_DATES)
        specific_dates_elem.click()

        current_date = datetime.date.today()

        first_day = current_date.replace(day=1)
        if n_prev_months > 1:
            for _ in range(n_prev_months):
                first_day = datetime.date(current_date.year, current_date.month - 1, 1)

        self.start_date, self.end_date = first_day, current_date
        keys = f'{self.start_date.strftime("%m/%d/%Y")}{Keys.TAB}{self.end_date.strftime("%m/%d/%Y")}{Keys.ENTER}'
        self.__browser.press_keys(None, keys)

    def __apply_section_filter(self, section=None):
        if isinstance(section, str):
            section = section.capitalize()

            section_div = self.__browser.find_element(SECTION_DIV)
            section_label = section_div.find_element(by=By.CLASS_NAME, value=SECTION_LABEL_CLASS)

            section_label.click()

            section_ul = section_div.find_element(by=By.CLASS_NAME, value=SECTION_UL_CLASS)
            section_li = section_ul.find_elements(by=By.CLASS_NAME, value=SECTION_LI_CLASS)[1:]

            for li in section_li:
                if section == re.sub(r'\d', '', li.text):
                    li.click()
                    break

            section_label.click()

    def __apply_filters(self, n_prev_months=0, section="World"):
        self.__apply_date_filter(n_prev_months=n_prev_months)
        self.__apply_section_filter(section=section)

    def __sort_results(self, order='newest'):
        self.__browser.select_from_list_by_value(SORT_BY, order)
        time.sleep(5)

    def __load_more(self):
        self.__browser.execute_javascript(f"window.scrollBy(0, 700);")
        self.result_items = []
        all_result_elems_list = []
        res_len = all_result_elems_list.__len__()

        current_date = datetime.date.today()
        show_more = True

        while show_more:
            result_elems_list = self.__browser.find_elements(SEARCH_RESULTS)

            for item_idx in range(res_len, result_elems_list.__len__()):
                item = result_elems_list[item_idx]
                res_date = item.find_element(by=By.CLASS_NAME, value=RESULT_DATE_CLASS).text
                if "ago" not in res_date:
                    # formatted_date = dateutil.parser.parse(res_date)
                    temp = res_date.split(' ')
                    month = temp[0][:3]
                    day = temp[1][:2]
                    month = datetime.datetime.strptime(month.capitalize(), '%b').month
                    formatted_date = datetime.date(current_date.year, month, int(day))

                    if not (self.start_date < formatted_date < self.end_date):
                        show_more = False
                        break

                res_item_obj = {'date': res_date,
                                'title': item.find_element(by=By.CLASS_NAME, value=RESULT_TITLE_CLASS).text}
                try:
                    res_item_obj['desc'] = item.find_element(by=By.CLASS_NAME, value=RESULT_DESC_CLASS).text
                except NoSuchElementException:
                    res_item_obj['desc'] = ''

                self.result_items.append(res_item_obj)
                all_result_elems_list.append(item)
            else:
                res_len = all_result_elems_list.__len__()
                self.__browser.click_button_when_visible(SHOW_MORE_BUTTON)
                self.__browser.execute_javascript(f"window.scrollBy(0, 1000);")

    def __create_excel_and_save_data_init(self):
        filename = 'NewsData2.xlsx'
        self.__files.create_workbook(filename)
        self.__files.set_active_worksheet(0)

        header = ['Date', 'Title', 'Description', 'Phrase Count', 'Contains any Amount']
        data = [header]

        for item in self.result_items:
            t_d_combined = f'{item["title"]} {item["desc"]}' if item["desc"] else item['title']

            phrase_count = re.findall(self.__phrase, t_d_combined, flags=re.IGNORECASE).__len__()

            amount_pattern = r'\$[0-9,]+(\.[0-9]+)?|\b[0-9]+ dollars\b|\b[0-9]+ USD\b'
            found = re.search(amount_pattern, t_d_combined)
            contains_any_amount = 'Yes' if found else 'No'

            new_row = [item['date'], item['title'], item['desc'], phrase_count, contains_any_amount]
            data.append(new_row)

        self.__files.append_rows_to_worksheet(data)
        self.__files.save_workbook()

    def run(self):
        self.__open_website()
        self.__close_modal()
        self.__search_phrase()
        self.__apply_filters(n_prev_months=2)
        self.__sort_results()
        self.__load_more()
        self.__create_excel_and_save_data_init()


phrase = 'pakistan'
nyt_robot = NYTimesNewsRobot(phrase=phrase)
nyt_robot.run()
