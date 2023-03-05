## Parse Excel to Modify XML Files
This Python script reads an Excel file and an XML file, replaces text in the XML file with data from the Excel file, and saves 5 modified copies of the XML file with names specified in the Excel file. The script uses the openpyxl library to read the Excel file and the xml.etree.ElementTree library to read and modify the XML file.

The script starts by trying to read the Excel file and the XML file. If there is an error opening either file, an error message is printed. If the files are successfully opened, the script iterates over 5 rows in the Excel file, reading the name for each modified copy of the XML file from the second column of each row. The script then enters a loop for each row that reads data from the third and fourth columns of the row and searches for text in the XML file that matches the data. If a match is found, the script replaces the text with the data from the fourth column of the row.

After all replacements have been made, the script saves each modified copy of the XML file using the name specified in the Excel file. To save the copies in their own folder called "result", the script uses the os module to create the "result" directory if it doesn't exist and navigate to it. The modified copies of the XML file are then saved in the "result" folder with the names specified in the Excel file. Finally, the script prints a message indicating that each copy has been saved.

In summary, this script is a useful tool for automating the modification of XML files based on data in an Excel file. It is flexible and can be easily modified to work with different Excel and XML files, making it a useful tool for a variety of data processing tasks.
## Installation
To run this script, you will need to install the following Python libraries:

- openpyxl: for reading Excel files
- xml.etree.ElementTree: for parsing XML files
- os: to create result directory

You can install these libraries using pip by running the following command:
```
pip install openpyxl xml.etree.ElementTree os
```
## Usage
- Open the Generationfile.xlsx Excel file and modify the data in the sheet as needed.
- Run the parse.py script using Python.
```
python parse.py
```
- The modified XML files will be saved in new directory called "result".

## Contributing
Contributions are welcome! If you would like to contribute to this project, please fork the repository and create a pull request with your changes.

## Copyright
Copyright (c) 2023 Maged Magdy. All rights reserved.

## Contact
For any inquiries or questions about the application, please contact the author at magedmagdy.engr@gmail.com.
