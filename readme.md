<!-- Run this command to activate environment -->
pipenv shell

<!-- Run this command to install all requirements -->
pipenv install

<!-- Run This command to run the server -->
uvicorn myapi:app --reload


<!-- If a pipenv is inside another env(possibly venv) follow these instructions: -->
Courtesy Notice:
Pipenv found itself running within a virtual environment,  so it will automatically use that environment, instead of  creating 
its own for any project. You can set
PIPENV_IGNORE_VIRTUALENVS=1 to force pipenv to ignore that environment and create  its own instead.
You can set PIPENV_VERBOSITY=-1 to suppress this warning.