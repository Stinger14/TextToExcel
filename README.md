# TextToExcel
Parses all text (*.txt) files in a directory into a single Excel file (.xsls) using the xlwt and xlrd modules through a simple PyQt5 GUI interface.

The content of the *.txt should be delimited with pipes (|) acting as column separator and each non-empty line is treated as a row for the files to be parsed properly.

REQUIREMENTS
- xlwt
- xlrd
- PyQt5

NOTE: Every file or folder in a NTFS volume has a owner. If a PermissionError pops you may need to change ownership of the directory your txt files are.

On windows open the command-line as Administrator and type: "takeown /f <foldername> /r /d y"

On Linux open terminal and type: "sudo chown -R username:group directory"

To install the python3 modules open the python interpreter through the cmd/terminal and type:

- pip install xlwt xlrd PyQt5
