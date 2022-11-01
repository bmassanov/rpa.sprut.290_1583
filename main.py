from datetime import datetime

from settings.logger import logger
from sources.__main__ import parse_excel, export_1583, export_290, fill_excel
from sources.tools import Sprut

start_time = datetime.now().replace(microsecond=0)
logger.info('=== START ===')
data = parse_excel()
for n, xl_list in enumerate(data):
    print(f'{n} form {len(data)}')
    obj, tz, name = xl_list[0], xl_list[1], xl_list[2]
    try:
        if int(tz) == 260:
            continue
        _1583 = export_1583(tz)
        print(str(_1583))
        _290 = export_290(obj)
        print(str(_290))
        fill_excel(_1583, _290, obj, name)
    except Exception as e:
        logger.warning(f'{obj}, {tz}, {name}')
        Sprut().quit()
        logger.exception(str(e))
logger.info(f'=== END === {(datetime.now().replace(microsecond=0) - start_time)}')
