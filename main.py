from datetime import datetime

from settings.logger import logger
from sources.__main__ import parse_excel, export_1583, export_290, fill_excel, check_serialisation, setto_serialisation

start_time = datetime.now().replace(microsecond=0)
logger.info('=== START ===')
for n, xl_list in enumerate(parse_excel()):
    try:
        obj, tz = xl_list[0], xl_list[1]
        _1583 = export_1583(tz)
        logger.info(str(_1583))
        _290 = export_290(obj)
        logger.info(str(_290))
        fill_excel(_1583, _290, n)
        setto_serialisation(obj)
    except Exception as e:
        logger.exception(str(e))
logger.info(f'=== END === {(datetime.now().replace(microsecond=0) - start_time)}')
