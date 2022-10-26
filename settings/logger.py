from logging import getLogger, Formatter, StreamHandler, INFO
from logging.handlers import RotatingFileHandler

from settings.paths import root_path, logs_path

c_format = '\n%(asctime)s|| %(levelname)s||(%(threadName)s->%(filename)s->%(funcName)s->%(lineno)d):\n* %(message)s'
c_formatter = Formatter(c_format)
c_handler = StreamHandler()
c_handler.setFormatter(c_formatter)
c_handler.setLevel(INFO)

f_path = logs_path.joinpath(f'{root_path.name}.log')
f_format = '%(asctime)s||%(levelname)s||(%(threadName)s->%(pathname)s->%(funcName)s->%(lineno)d): %(message)s'
f_formatter = Formatter(f_format)
f_handler = RotatingFileHandler(f_path, maxBytes=2 * 1024 * 1024, backupCount=100, encoding="utf-8")
f_handler.setFormatter(f_formatter)
f_handler.setLevel(INFO)

logger = getLogger(root_path.name)
logger.setLevel(INFO)
logger.addHandler(c_handler)
logger.addHandler(f_handler)
