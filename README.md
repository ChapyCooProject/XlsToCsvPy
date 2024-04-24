## Excel-CSV Conversion (Japanese Edition)

### Development Environment

   * Python 3.8
   * Windows 10 64-bit


### Introduction

- This software is designed solely to convert Excel files to CSV files.
- The software "ExcelToCsvPy.exe" operates as a standalone executable.
- It supports both XLS and XLSX file formats.
- It can handle password-protected Excel files.
- It does not support multiple sheets.
- It is only compatible with 64-bit Windows.
- No additional components such as EXCEL or ACCESS are required.

### Method 1

1. Launch the executable file (ExcelToCsvPy.exe).
2. A dialog will prompt you to select the Excel file.
3. Select the Excel file you wish to convert and click the open button.
4. If prompted, input the password if necessary.
5. A CSV file with the same name will be created in the same location as the Excel file.

### Method 2

1. Drag and drop the Excel file you want to convert onto the executable file.
2. If prompted, input the password if necessary.
3. A CSV file with the same name will be created in the same location as the Excel file.

### Method 3 (Common)

1. Launch the command prompt and navigate to the directory where the executable file is located.
2. Set the full path of the Excel file you want to convert as an argument and launch the executable file.
   ```cmd
   ExcelToCsvPy C:\sampleXLSX.xlsx
   ```
3. If prompted, input the password if necessary.
4. A CSV file with the same name will be created in the same location as the Excel file.

### About Configuration Files (.json)

- Creating a configuration file (.json) with the same name as the Excel file you want to convert allows for detailed customization.
- When using a configuration file, place it in the same location as the Excel file you want to convert.

### About Configuration Parameters

- The parameters available in the configuration file are as follows:
  - `sheet_name`: Specifies the name of the sheet to be read. (Not applicable to multiple sheets)
  - `excel_password`: Sets the password for password-protected Excel files.
  - `has_header`: Specifies whether to output the header row to the CSV file.
    - `True`: Output
    - `False`: Do not output
  - `has_index`: Specifies whether to output the row number column to the CSV file.
    - `True`: Output
    - `False`: Do not output
  - `encoding`: Specifies the character encoding to be used for the CSV file.
    - `utf_8`: UTF-8
    - `shift_jis`: Shift-JIS
    - `Other`: Other character encodings compliant with Python
  - `sep_char`: Specifies the delimiter when outputting to the CSV file.
    - `,`: Comma-separated
    - `\t`: Tab-separated
    - `Other`: Other strings compliant with Python
  - `write_mode`: CSV file output mode.
    - `w`: New (existing files will be overwritten)
    - `x`: New (existing files will not be overwritten)
    - `a`: Append
  - `quoting`: Specifies the quoting character when outputting to the CSV file.
    - `0`: QUOTE_MINIMAL (Quotes fields containing special characters such as delimiter, quotation, and newline.)
    - `1`: QUOTE_ALL (Quote all fields.)
    - `2`: QUOTE_NONNUMERIC (Quote all non-numeric fields.)
    - `3`: QUOTE_NONE (Do not quote any fields. Delimiters in values are escaped with the escape character set.)
  - `column_mapping`: Sets the definition for each column.
    - `First number`: Column number
    - `csv_col_name`: Column header name when outputting to the CSV file.
    - `data_type`: Considers the data type when outputting to the CSV file.
      - `Integer`: Output as numerical (integer) data.
      - `Decimal`: Output as numerical (decimal) data.
      - `String`: Output as string data.
      - `Date`: Convert date type data to the specified format and output as string data.
        > fmt: Date-time notation (notation on CSV file)※The notation method complies with Python.
      - `Time`:
        > fmt_from: Notation before conversion (notation on Excel file)※The notation method complies with Python.  
        > fmt_to: Notation after conversion (notation on CSV file)※The notation method complies with Python.

        `※Python datetime notation`
        ```
        %Y: Represents the year (4 digits).
        %y: Represents the year (2 digits).
        %m: Represents the month.
        %d: Represents the day.
        %H: Represents the hour (24-hour clock).
        %I: Represents the hour (12-hour clock).
        %p: Represents AM or PM.
        %M: Represents the minute.
        %S: Represents the second.
        ```

### Downloading the Executable File

  - The executable file can be downloaded from the following URL.  
    [Download](https://drive.google.com/drive/folders/1b3sa8OYJcE8Brg22wiagkk7OABDlZ1bb?usp=sharing)

