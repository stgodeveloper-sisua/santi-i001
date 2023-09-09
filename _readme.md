# Sisua python framework
- Created by: vicente.diaz & juanpablo.mena & guillermo.konig
- Creation date: 2023-01-16
- Last updated:  2023-08-18

## Objetive
This framework allows you to develop a python process in a structured and robust way.

### Points to consider when using the framework:
* Follow the [PEP8](https://peps.python.org/pep-0008/) and [PEP20](https://peps.python.org/pep-0020/) python best practice guides.
* Create a virtual environment before starting development using [venv](https://docs.python.org/3/library/venv.html#module-venv) library.
* **Do not modify** the `robot.py` and `modules/_fmw/fmw_utils.py` files.

### Starting a new project
1. Clone the python framework in the desired directory and change the folder name to the project name.
2. Open the project with VSCode application.
3. **Create a virtual environment** using the following commands:
   >  python -m venv _venv
   > _venv\Scripts\activate.bat

    If the following error message appears: "*`activate.bat` cannot be loaded because running scripts is disabled on this system*", it can be solved using the options 1 or 2 of the website
    [Fix UnauthorizedAccess](https://support.enthought.com/hc/en-us/articles/360058403072-Windows-error-activate-ps1-cannot-be-loaded-because-running-scripts-is-disabled-UnauthorizedAccess-).

    Execute the following command in the project directory to check if the virtual environment was created successfully:
    > pip freeze

    If the the output is empty, the venv setup was successful and it is ready to install the project `requirements.txt` (if exists) with the command:
    > pip install -r requirements.txt

    To create/update the `requirements.txt` file, execute the command:
    > pip freeze > requirements.txt

    To deactivate the virtual environment, just execute the following command:
    > deactivate
4. Select the virtual environment python interpreter in the Command Palette (Select Interpreter) or clicking in the right bottom panel.
5. Run the `robot.py` script to check that the setup was successful.
6. Import [Sisua python libraries](https://github.com/sisua/sisuacl_python_libraries) if needed. 
7. Backup your project in a Sisua Github repository using the following commands:
    > git clone https://github.com/sisua/sisuacl_python_framework

**Now you are ready to start developing your project!**
