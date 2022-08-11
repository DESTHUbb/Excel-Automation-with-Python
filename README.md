# Excel-Automation-with-Python
 This can be useful if you need to do repetitive tasks with students, grades, daily tasks...
 
 ![excel-automation-python](https://user-images.githubusercontent.com/90658763/184173123-9686bbb2-eb0e-4389-ae91-ceb691dc45f6.jpg)

## Instructions for use

* 1: Installing openpyxl

`	pip install openpyxl` 

* 2: trial installation

```python
from openpyxl
```

* 3: Load an existing workbook
```python
from openpyxl import Workbook, load_workbook

wb = load_workbook('C:/Users/..../Grades.xlsx')
```

* 4: Access to worksheets
```python
ws = wb.active
print(ws)
```

* 5: Accessing Cell Values
```python
WS['A2'].value = "Test"
```


* 6: Save workbooks
```python
WB.save('Grades.xlsx')
```
![image](https://user-images.githubusercontent.com/90658763/184167940-43f54541-c545-4314-a121-e4d646fd314b.png)
![image](https://user-images.githubusercontent.com/90658763/184168042-b0376f72-3b7f-44d7-80d8-75f2e8b3603c.png)

* 7: Create, list, and change sheets
```python
from openpyxl import Workbook, load_workbook
wb = load_workbook('Grades.xlsx')
wb.create:sheet("Test")

print(wb.sheetnames)
```

* 8: Add rows
```python
from openpyxl import Workbook, load_workbook

wb = Workbook()
ws = wb.active
ws.title = "Data"

ws.append(['Joe', 'Is', 'Great', '!'])
ws.append(['Joe', 'Is', 'Great', '!'])
ws.append(['Joe', 'Is', 'Great', '!'])
ws.append(['Joe', 'Is', 'Great', '!'])
ws.append(["end"])

wb.save('joe.xlsx')
```
Open file joe.xlsx

![image](https://user-images.githubusercontent.com/90658763/184169507-5c976ac3-173c-48c9-9ce4-623b36a3c2fc.png)


* 9: Access to multiple cells
```python
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter


wb = load_workbook('joe.xlsx')
ws = wb.active

for row in range(1, 11):
    for col in range(1, 5):
        char = get_column_letter(col)
        ws[char + str(row)] = char + str(row)
       
wb.save('joe.xlsx')
```
Open file joe.xlsx

![image](https://user-images.githubusercontent.com/90658763/184170099-7bee680e-68cc-469b-9912-4b5463c05712.png)


* 10: Merge cells
```python
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

wb = load_workbook('joe.xlsx')
ws = wb.active

ws.merge_cells("A1:D2")

wb.save('joe.xlsx')
```
Open file joe.xlsx

![image](https://user-images.githubusercontent.com/90658763/184170335-91a3ac7a-ec78-4975-9587-af738c5daa03.png)

* 11: Insert and delete rows
```python
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter


wb = load_workbook('joe.xlsx')
ws = wb.active

ws.insert_rows(7)
ws.insert_rows(7)

wb.save('joe.xlsx')
```
Open file joe.xlsx

![image](https://user-images.githubusercontent.com/90658763/184171080-4247e639-dce5-448c-bd2a-249d5db3f06a.png)

```python
ws.delete_rows(7)
```
Open file joe.xlsx

![image](https://user-images.githubusercontent.com/90658763/184171519-19ec994d-360e-44b6-8f9c-4ae2de9b4222.png)


* 12: Inserting and Deleting Columns
```python

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

wb = load_workbook('joe.xlsx')
ws = wb.active

ws.insert_cols(2)

wb.save('joe.xlsx')
```
Open file joe.xlsx

![image](https://user-images.githubusercontent.com/90658763/184171945-a5081f65-335b-4f46-ab51-b260003b7513.png)

```python
ws.delete_cols(2)
```
![image](https://user-images.githubusercontent.com/90658763/184172116-49866d92-3347-4f0a-8dfe-24cdeba54b8c.png)

* 13: Copy and move cells
```python
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

wb = load_workbook('joe.xlsx')
ws = wb.active

ws.move_range("C1:D11", rows=2, cols=2)

wb.save('joe.xlsx')
```
Open file joe.xlsx

![image](https://user-images.githubusercontent.com/90658763/184172443-d1111cc9-73bb-4817-b95a-03d87a416f40.png)

