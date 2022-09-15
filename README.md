# Simple example to handle Excel files with python

Just a prove of concept to run and modify Excel files programmatically.

## run-excel.py

Some simple commands to edit the file and run a simple macro.

`python run-excel.py __PROJECT_PATH__/dummy.xlsm`

## run-excel-server.py

Webserver to trigger the process by calling endpoints.

- Start: `uvicorn run-excel-server:app --reload`

- URL: `http://localhost:3000`