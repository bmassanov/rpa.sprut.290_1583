from datetime import datetime

from settings.logger import logger
from settings.paths import downloads_path, codes_path, report_path, serialisation_path
from settings.settings import report_month, start_date, end_date
from sources.rpamini import App, Json
from sources.tools import Sprut, Xls, Excel

# date = datetime(datetime.today().year, datetime.today().month, 1) - timedelta(days=1)
# date = datetime(datetime.today().year, datetime.today().month, 28)

serialisation_file = serialisation_path.joinpath(f'{datetime.today().year}.{report_month}.json')
MONTHS = [
    'Январь',
    'Февраль',
    'Март',
    'Апрель',
    'Май',
    'Июнь',
    'Июль',
    'Август',
    'Сентябрь',
    'Октябрь',
    'Ноябрь',
    'Декабрь'
]
VALUES = [
    'Всего',
    'Из них продовольственные товары',
    'Товарные запасы на конец отчетного месяца'
]


def check_serialisation(name):
    if serialisation_file.is_file():
        if name in Json.read(serialisation_file):
            return True
        else:
            return False
    else:
        return False


def setto_serialisation(name):
    if serialisation_file.is_file():
        data = Json.read(serialisation_file)
    else:
        data = list()
    data.append(name)
    Json.write(serialisation_file, data)


def export_1583(object_code: int):
    value_ = f'{report_month}_1583_{object_code}.xls'
    path_ = downloads_path.joinpath(value_)
    if path_.is_file():
        return path_
    sprut = Sprut()
    sprut.authorize().open_module('Отчеты')
    sprut.search(1, 0, 2202)
    sprut.get_pane(1).type_keys(sprut.Keys.F9)
    sprut.set_input(6, start_date)
    sprut.set_input(4, end_date)
    sprut.set_select(3, 'Весь период')
    sprut.set_multiselect(2, [(0, 0, object_code, True)])
    selector_ = [{"class_name": "TfrmParams", "index": 0},
                 {"control_type": "Edit", "index": 1}]
    App.find_element(selector_).type_keys(App.Keys.DOWN + App.Keys.SPACE)
    App.find_element(selector_).type_keys(str(App.Keys.DOWN * 2) + '^{SPACE}' + App.Keys.ENTER)
    selector_ = [{"class_name": "TfrmParams", "index": 0},
                 {"title": "Ввод", "control_type": "Button", "index": 0}]
    sprut.find_element(selector_).click_input()

    selector_ = [{"control_type": "Window", "class_name": "XLMAIN", "index": 0}]
    App.find_element(selector_, timeout=3600).click_input()
    App.find_element(selector_).type_keys('{F12}')

    selector_ = [{"control_type": "Window", "class_name": "XLMAIN", "index": 0},
                 {"control_type": "Edit", "class_name": "Edit", "index": 0}]
    element_ = App.find_element(selector_)
    element_.type_keys(path_.__str__() + App.Keys.ENTER)

    selector_ = [{"control_type": "Window", "class_name": "XLMAIN", "index": 0},
                 {"title": "Да", "control_type": "Button", "class_name": "CCPushButton", "index": 0}]
    if App.wait_element(selector_, timeout=2):
        App.find_element(selector_).click_input()

    selector_ = [{"control_type": "Window", "class_name": "XLMAIN", "index": 0}]
    App.find_element(selector_).close()
    sprut.quit()
    return path_


def export_290(object_code: int):
    value_ = f'{report_month}_290_{object_code}.xls'
    path_ = downloads_path.joinpath(value_)
    if path_.is_file():
        return path_
    sprut = Sprut()
    sprut.authorize().open_module('Отчеты')
    sprut.search(1, 0, 2360)
    sprut.get_pane(1).type_keys(sprut.Keys.F9)
    sprut.set_multiselect(11, [(5, 2, object_code, True)])
    sprut.set_select(10, 'Нет')
    sprut.set_input(8, start_date)
    sprut.set_input(6, end_date)
    sprut.set_select(5, 'Все товары [2]')
    selector_ = [{"class_name": "TfrmParams", "index": 0},
                 {"title": "Ввод", "control_type": "Button", "index": 0}]
    sprut.find_element(selector_).click_input()

    selector_ = [{"control_type": "Window", "class_name": "XLMAIN", "index": 0}]
    App.find_element(selector_, timeout=3600).click_input()
    App.find_element(selector_).type_keys('{F12}')

    selector_ = [{"control_type": "Window", "class_name": "XLMAIN", "index": 0},
                 {"control_type": "Edit", "class_name": "Edit", "index": 0}]
    element_ = App.find_element(selector_)
    element_.type_keys(path_.__str__() + App.Keys.ENTER)

    selector_ = [{"control_type": "Window", "class_name": "XLMAIN", "index": 0},
                 {"title": "Да", "control_type": "Button", "class_name": "CCPushButton", "index": 0}]
    if App.wait_element(selector_, timeout=2):
        App.find_element(selector_).click_input()

    selector_ = [{"control_type": "Window", "class_name": "XLMAIN", "index": 0}]
    App.find_element(selector_).close()
    sprut.quit()
    return path_


def parse_excel():
    xlsx = Excel(codes_path)
    rows = [row for row in list(xlsx.ws.values)]
    list_ = [(int(row[4]), int(row[3])) for row in rows if 'торговый зал' in str(row[2]).lower()]
    return list_


def fill_excel(path_1583, path_290, branch_index):
    data = dict()
    xls = Xls(path_1583.__str__())
    result = {
        'month': MONTHS[report_month - 1],
        'branch': xls.get(xls.find('Итого:')[0][0] - 1, xls.find('Компания')[0][1]),
        'index': branch_index,
        'values': None
    }
    data[VALUES[0]] = int(xls.get(xls.find('Итого:')[0][0], xls.find('Оборот, тг без НДС')[0][1]) / 1000)
    data[VALUES[1]] = int(data["Всего"] - data["Всего"] * 0.15)
    xls = Xls(path_290.__str__())
    _290_1 = xls.get(xls.find('ИТОГО:')[0][0], xls.find('Сумма проблемного прихода')[0][1])
    _290_2 = xls.get(xls.find('ИТОГО:')[0][0], xls.find('Товарный остаток на конец, тг')[0][1])
    data[VALUES[2]] = int((_290_2 - _290_1) / 1000)
    result['values'] = data

    xlsx = Excel(report_path, str(datetime.today().year))
    xlsx.init(MONTHS)
    print(result)
    xlsx.fill(result)


if __name__ == '__main__':
    start_time = datetime.now().replace(microsecond=0)
    logger.info('=== START ===')
    for n, xl_list in enumerate(parse_excel()):
        obj, tz = xl_list[0], xl_list[1]
        if check_serialisation(obj):
            continue
        _1583 = r"C:\Users\ASSANOV.B\rpa.sprut.290_1583\downloads\1583_945.xls"
        _290 = r"C:\Users\ASSANOV.B\rpa.sprut.290_1583\downloads\290_815.xls"
        fill_excel(_1583, _290, n)
        setto_serialisation(obj)
    logger.info(f'=== END === {(datetime.now().replace(microsecond=0) - start_time)}')
