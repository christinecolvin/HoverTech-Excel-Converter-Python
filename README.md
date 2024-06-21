# HoverTech-Excel-Converter
An interactive graphical user interface that allows the user to choose a Microsoft Excel Workbook (.XLSX) file as well as a save file name & space. The excel file is then converted to filter, sort, and run calculations and exported into the chosen save file & space. 

Project designed for *HoverTech International* by Christine Colvin, intern.

## System Architecture
![HoverTech Diagram.png](https://github.com/christinecolvin/HoverTech-Excel-Converter-Python/blob/main/HoverTech%20Diagram.png)

### Contributors:
- [Christine Colvin](https://github.com/christinecolvin)

# Local Installation Tutorial

### 1. Clone the repo
Press the green **<> Code** button to gain a link to clone the repository.

Then, open your preferred [Command Line Interface](https://en.wikipedia.org/wiki/Command-line_interface#:~:text=A%20command%2Dline%20interface%20\(CLI,interface%20available%20with%20punched%20cards.) and clone the repository with the following command:

```
git clone https://github.com/christinecolvin/HoverTech-Excel-Converter-Python.git 
```
### 2. Dependencies 
If any dependencies are not installed on your system, use this command with the blank replacing the dependency not in access: 
```
pip install ___
```
### 3. Run `OSTkinter.py` 
This will generate the Tkinter GUI and allow direct interaction. 

Please heed the message boxes that may pop up during your use. 

# Building an .exe file 

### 1. Clone the repo
Press the green **<> Code** button to gain a link to clone the repository.

Then, open your preferred [Command Line Interface](https://en.wikipedia.org/wiki/Command-line_interface#:~:text=A%20command%2Dline%20interface%20\(CLI,interface%20available%20with%20punched%20cards.) and clone the repository with the following command:

```
git clone https://github.com/christinecolvin/HoverTech-Excel-Converter-Python.git
```
### 2. Using .Py to .Exe
The simplest way to create a one-file executable for a desktop application is to install it using [PyPi](https://pypi.org/). and [Auto-.Py-to-.Exe](https://github.com/brentvollebregt/auto-py-to-exe)

Open your preferred [Command Line Interface](https://en.wikipedia.org/wiki/Command-line_interface#:~:text=A%20command%2Dline%20interface%20\(CLI,interface%20available%20with%20punched%20cards.) and copy this command:
```
pip install auto-py-to-exe
```
To then use the .py to .exe interface, run this command:

```
auto-py-to-exe
```
If this does not prompt you with the website, try this command:
```
python -m auto_py_to_exe
```
### 3. Creating the .exe file 
Once [Auto-.Py-to-.Exe](https://github.com/brentvollebregt/auto-py-to-exe) is installed and you're now interacting with the interface,

- Select the `OSTkinter.py` file, which is located in the `Tkinter-GUI` folder for the script location.
- Select  `One File`
- Under Additional Files, select `Add files` and navigate to the `Tkinter-GUI` folder then select the `forest-dark.tcl` file and `forest-light.tcl`
- Under Additional Files, select `Add folders` and navigate to the `Tkinter-GUI` folder then select the `forest-dark` folder and `forest-light` folder
- Its important to name the .exe file. Under Advanced, give the file a preferred name.
- Once these steps are completed, we can now press the blue `covert` button.
- This may take a moment. After it is finished, press `open output folder` and the .exe file should open on a click.

