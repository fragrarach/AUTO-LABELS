from quatro import get_parent
from subprocess import call


parent_path = get_parent()
parent_name = parent_path.split("\\")[-1]

venv_path = f'{parent_path}\\venv\\Scripts\\python.exe'
script_path = f'{parent_path}\\{parent_name}.py'

command = f'"{venv_path}" "{script_path}"'

call(command)
