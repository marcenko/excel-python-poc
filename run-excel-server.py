from fastapi import FastAPI
import sys
import os
import subprocess

app = FastAPI()

@app.get('/')
def call_excel_macro():
    subprocess.Popen([sys.executable, 'run-excel.py'])
    return 'Excel done'