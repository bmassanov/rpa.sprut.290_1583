from pathlib import Path
from shutil import copy

root_path = list(Path(__file__).parents)[1]

project_name = root_path.name
project_path = Path.home().joinpath(project_name)
project_path.mkdir(exist_ok=True)

logs_path = project_path.joinpath('logs')
logs_path.mkdir(exist_ok=True)

downloads_path = project_path.joinpath('downloads')
downloads_path.mkdir(exist_ok=True)

serialisation_path = project_path.joinpath('serialisation')
serialisation_path.mkdir(exist_ok=True)

app_path = r'C:\SPRUT\Modules3.5\sprut.exe'
template_path = root_path.joinpath('settings\\codes.xlsx')
codes_path = project_path.joinpath('codes.xlsx')
if not codes_path.is_file():
    copy(template_path.__str__(), codes_path.__str__())
report_path = project_path.joinpath('report.xlsx')
