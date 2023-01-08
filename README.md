# Excel And CSV File Converter GUI
### About
##
A standalone GUI app created using Python to convert Excel files to CSV and vice-versa. Created using [PySimpleGUI](https://www.pysimplegui.org/en/latest/), [Pandas](https://pandas.pydata.org/), [openpyxl](https://openpyxl.readthedocs.io/en/stable/), and [psgcompiler](https://pypi.org/project/psgcompiler/).

![GUI_example](https://user-images.githubusercontent.com/31321037/211216398-db64c69d-bdbd-4ac7-b56f-9a9a07bb0f1f.PNG)

##
### Compile
##

Download and launch the [.exe file](https://github.com/chrisdmancuso/excel-csv-gui/blob/main/ExcelGUI.exe). 

Optionally: follow along to compile from source code, or use your preferred py to exe method.

0) Download [main.py](https://github.com/chrisdmancuso/excel-csv-gui/blob/main/main.py) and [config.ini](https://github.com/chrisdmancuso/excel-csv-gui/blob/main/config.ini) to your local machine

1) [Ensure pip is installed](https://docs.python.org/3/library/ensurepip.html)
2) Using Command Line, create a Python virtual environment
```
python -m venv your-folder-name-here
```

3) Navigate to the created folder, and activate the virtual environment
```
cd your-folder-name-here
```
```
.\scripts\activate
```
If successfully activated, the virtual environment will appear in parentheses next to the working directory like so

![install_env](https://user-images.githubusercontent.com/31321037/211218002-8a4166dc-aafe-4801-b58e-170d472a2942.PNG)

4) Run pip install "package-name" for the necessary packages (pandas, openpyxl, PySimpleGUI, psgcompiler)
```
pip install pandas openpyxl PySimpleGUI psgcompiler
```

5) Run psgcompiler within the virtual environment

![psg_env](https://user-images.githubusercontent.com/31321037/211218070-0c798f42-63ae-4cce-a999-b9392a044fee.PNG)

6) Select main.py as your Python script, and convert

![psg_compiler](https://user-images.githubusercontent.com/31321037/211218091-de90d067-4cf8-4a19-91af-213b19510a5d.PNG)

7) Deactivate your virtual environment

![deactivate_env](https://user-images.githubusercontent.com/31321037/211218110-ffc135d1-7397-4292-9852-450a7cedbb53.PNG)

8) Launch your new .exe!
##
### Usage
##

1) Select a .xlsx or .csv file to convert.
2) Select the output folder for your new file
3) Optional: Give your new file a name. Defaults to the name of the converting file.
4) Convert!

![GUI_example_steps](https://user-images.githubusercontent.com/31321037/211217130-2606c9bb-6e78-4b4d-9ae9-9ad54b7c62da.png)
