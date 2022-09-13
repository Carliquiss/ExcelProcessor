# ExcelProcessor
Python library to work with Excel files (.CSV, .XLS and .XLSX). Intended for scientific analysis

---

# ⓘ Progress: 
## 📖 Read from file extensions: 
- .xls files: ✓
- .xlsx files: ✓
- .csv files: ✓

## ✏️ Write to file extensions (pending): 
- .xls files: X
- .xlsx files: X
- .csv files: X

---

# ⚙️ Usage

The main use of the script is divided into 2 options:

- Read a folder with **multiple** files 
- Read a **single** file.

<details>
    <summary>
        <h2> 📁 Read a folder with multiple files</h2>
    </summary>
   
➡️ To read all files in a folder and store them as objects in an array: 

```
AllFiles = ExcelFolder("../Stuff")
```

➡️ To check that the files have been read correctly do the following print:

```
AllFiles.PrintGoodFiles()
```

Output: 
```
Filename - ../Stuff/Test File 1.csv
	Columns: 2
	Rows: 3

Filename - ../Stuff/Test File 2.xls
	Columns: 16
	Rows: 124

Filename - ../Stuff/Test File 3.xlsx
	Columns: 16
	Rows: 124

Filename - ../Stuff/Test File 4.csv
	Columns: 16
	Rows: 124

Filename - ../Stuff/Test File 5.xls
	Columns: 16
	Rows: 124
```

➡️ In addition, you can check if a file could not be opened:  
```
AllFiles.PrintBadFiles()
```

Output:
```
[X] Bad file: ../Stuff/NewTestFile1.csv

[X] Bad file: ../Stuff/Test File (2).csv

[X] Bad file: ../Stuff/TestFile1.csv
```

### 💡 For CSV files you can set the delimiter character: 
```
AllFiles = ExcelFolder("../ExcelsFolder", csvDelimiter=',')
```
</details>





<details>
    <summary>
        <h2> 📄 Read a single file</h2>
    </summary>

Documentation pending

</details>



