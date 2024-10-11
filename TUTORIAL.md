# **Microsoft Office**

MicroSoft Office is bunch of different apps or tools that are used in offices and Corporates # It was announced by Bill Gates in 1988

- **MS EXCEL**
- **POWERPOINT**
- **WORD**

**MS EXCEL**

MS EXCEL is a spreadsheet software created by Microsoft that is used to organize numbers and data in a sheet.

**PURPOSES**

- Data Entry
- Data management
- Accounting
- Financial Analysis
- Charting and reporting
- Programming
- Business Analysis
- people Management
- Managing Operation
- Performance Reporting
- Office Administration
- Performance Reporting
- Strategic Analysis
- Project Management

## Excel Terminology Simplified

- **Ribbon**: The toolbar at the top of Excel with all the options and tools (like File, Home, Insert, etc.).
- **Name Box**: The small box above the worksheet that shows the name of the selected cell or range (e.g., A1).
- **Formula Bar**: A bar above the worksheet where you can view and edit the formula or value of the selected cell.
- **Cell**: A single box in the Excel grid where you can enter data (e.g., A1, B2, etc.).
- **Row**: A horizontal line of cells, numbered on the left side (e.g., 1, 2, 3...).
- **Column**: A vertical line of cells, labeled at the top with letters (e.g., A, B, C...).
- **Zooming**: You can zoom in or out of the sheet.
  - Use `Alt + V` then `Z` to open zoom options.
  - Use `Ctrl + +` to zoom in or `Ctrl + -` to zoom out.
- **Last Column**: The farthest right column in Excel. Use `Ctrl + Right Arrow` to reach it. Excel has 16,384 columns and 1,048,576 rows.
- **Cell Formula**: To do a calculation in a cell, start with an equals sign `=`. For example:
  ```excel
  =60+40
  ```
- **Range**: A group of cells. For example, the range

  ```excel
  C3:F9
  ```

  includes all the cells from C3 to F9.

## Functions, Formula , Shortcuts

- `=SQRT(16)`

  - to calculate the square root of 16.

- `=C9+D3` (in cell D4)
  - In D4, enter `=` then select cell `C9`, enter `+`, then select `D3`, and press Enter.
  - You can do the same with subtraction (`-`), multiplication (`*`), and division (`/`).
- `=POWER(2,4)`
  - to calculate 2 raised to power 4
- `Inceasing Decreasing decimal digit `
  - In the Number section of Hom tab ,you will get to see two button `<-0.00` `->0.00`
    the first one is used to to increase and other one used to decrease decimal digit
- `Inserting New column` :
  - to insert new column click right button on the column index and then in the dropdown menu, click on insert,it will insert a new column to the left of the corresponding column .Same for row.Same thing can be done by right cicking on cell .It will provide you with several option.do as you wish.
- `Merge & Center` :
  - It is present in the Alignment section of home tab.It helps to convert multiple cell to one single cell.
- `Increasing title cell to fit the full text`:
  - hover over column index edge and double left click
  - select the columns then press`alt > h > o > i`
- `Entering Serial Number shortcut:`
  - Enter the first cell value as 1 and then the second cell value (e.g., 2).
  - Select both cells and drag the fill handle to extend the sequence as needed.
  - The sequence will follow the gap between the two selected cells. For example, if the gap between 1 and 2 is 1, the sequence will be 1, 2, 3, 4, 5, 6 and so on.
  - If the gap is 4 (e.g., 1 and 5), the sequence will be 1, 5, 9, 13, 17, 21, and so on.
  - Another way is to press Ctrl while dragging the fill handle. After dragging, release the mouse.
- `Data Entry :`
  - For faster data entry ,we need to follow the process.We have select the area in which we need to enter the data,then keep entering data and press enter to go next cell ,this way it can get faster.
- `=SUM(RANGE) ; =MAX(RANGE) ; =MIN(RANGE) ; =AVERAGE(RANGE)`
  - Click on sum in the editing section of home tab,and then select the cells you want to sum and then press enter.same for max,min,avg ,etc
  - `alt + =`
  - to do the same with following cells ,drag or double click the fill handle.
  - `% `:
    - enter `=` and then the press the total cell,enter `/` to total marks and `*` to `100`.

## Conditional Formatting

- select the region where you want to apply the operation
- Go to the `Home` tab -> click on `Styles` -> select `Highlight Cells Rules` -> choose `Greater Than`. This will apply a different color to numbers greater than the specified value..You can contrast with border,colors
- to remove the highlight ,select `clear rule`

### 1. Highlight Cells Rules

- **Greater Than**: Highlight cells greater than a specified value.
- **Less Than**: Highlight cells less than a specified value.
- **Between**: Highlight cells between two specified values.
- **Equal To**: Highlight cells equal to a specified value.
- **Text that Contains**: Highlight cells containing specific text.
- **A Date Occurring**: Highlight cells with dates in a certain range (e.g., yesterday, today).
- **Duplicate Values**: Highlight duplicate or unique values.

### 2. Top/Bottom Rules

- **Top 10 Items**: Highlight the top 10 (or specified number) highest values.
- **Top 10%**: Highlight the top 10% of values.
- **Bottom 10 Items**: Highlight the bottom 10 (or specified number) lowest values.
- **Bottom 10%**: Highlight the bottom 10% of values.
- **Above Average**: Highlight values that are above the average.
- **Below Average**: Highlight values that are below the average.

### 3. Data Bars

- Use bars to visually represent the values in cells. The length of the bar reflects the cell value.

### 4. Color Scales

- Apply a range of colors based on values, creating a gradient from low to high values.

### 5. Icon Sets

- Use icons like arrows, shapes, or traffic lights to represent relative cell values.

### 6. New Rule

- Create a custom rule using formulas for more advanced conditional formatting.

### 7. Clear Rules

- **Clear Rules from Selected Cells**: Remove formatting from selected cells.
- **Clear Rules from Entire Sheet**: Remove all formatting from the worksheet.

## Hiding the number in an specfied region

- `Home`->`Number`->`General drop down menu`->enter `;`

## IF / IF-AND / IF-OR / IF-IF :

**IF**

```excel
  =IF(LOGICAL TEST,[VALUE IF TRUE],[VALUE IF FALSE])
```

**IF-AND**

```excel
  =IF(AND(C3>40,D3>40,E3>40,F3>40,G3>40),"PASS","FAIL")
```

**IF-OR**

```excel
  =IF(OR(C4<40,D4<40,E4<40,F4<40,G4<40),"FAIL","PASS")
```

**IF-IF**

```excel
=IF(M3="FAIL","F",IF(L3>=75,"A+",IF(L3>=60,"A",IF(L3>=50,"B",IF(L3>40,"C","D")))))

```

## FORMAT AS TABLE:

`HOME` -> `STYLES` -> `FORMAT AS TABLE` -> choose format

## Filtering Sorting

- `alt > A > T`
- `home` -> `editing` -> `sort and filter`
- `data`->`sort and filter`
- **Equals**: Displays rows where the cell content exactly matches the specified text.
- **Does Not Equal**: Displays rows where the cell content does not exactly match the specified text.
- **Begins With**: Filters rows where the cell content starts with the specified text.
- **Ends With**: Filters rows where the cell content ends with the specified text.
- **Contains**: Displays rows where the cell content includes the specified text anywhere within it.
- **Does Not Contain**: Displays rows where the cell content does not include the specified text.
- **Custom Filter**: Allows combining multiple filter criteria, such as filtering for text that contains or begins with specific characters.

## Charts - Bar,Column,graph,PieChart

- `Insert Tab` -> `Charts section` select your desired charts ,and to customize it ,click on the chart and then **chart design** will pop up in the ribbon,then customize it accordingly.
  - Add title
  - omit vertical heading or horizontal heading
  - change color combination..etc

##

### Format Painter

- `Home tab`-> `clipboard`->format painter(looks like paint brush)
  Click on the cell whose style you want to copy, then double-click the Format Painter. After that, select the cell or range of cells where you want to apply the copied format.

### Removing border

- select the region where you want remove the border,then click on the windows shape icon ..You will not see the effect right away,once you fill the region with color ,you will get to see there is no border

### Text color:

- An Icon looks like `A`,clicking on it will change the color of selected cells' text.

### changing the row-width :

- Select the region and then hover at row line ,Dragging it will increase the width of all the row's row-width
- In cells section of home tab,format option has various option
  - row/column height
  - autofit
  - hide/unhide

### Wrapping the text

- wrap-text:present in alignment section of home tab.It wraps the text in the cell.
- `alt + enter`

### Printing excel sheet in single page:

- Print: `file -> print -> show preview` or `ctrl + p`
- Grid line enabling: `page layout -> sheet option -> tick the **print** check box`

### Inserting bullet point:

- `Insert Tab -> Symbol -> Symbol -> in verdana font `
- `=CHAR(149)`-> enter
- `=CHAR(149)&""&A1`

### Grouping ungrouping

- `data tab`->`Outline`->`Grouping`
- It will create a separation between regions based on the options you select, such as row or column.
- to Ungropup ,select the region you have grouped earlier and then click on `ungroup`

### Adding logo in Excel

- `Insert`-> `Text`->`Header and footer`
- You can insert picture,You just have to navigate to the header & footer tab(newly popped),clicked on Picture.
- Format picture in newly popped header & footer tab.
  update height
- To change the position, press enter before &picture.
- To reduce the transpancy, navigate to `picture format` and then `picture`,in automatic option, select `washout`.
- To dd the text as a logo, remove &picture and enter the logo text.Customize the text as needed.

### To hide a row:

- right click on te row heading,select `hide` option in the menu"
- To select non-consecutive rows, hold down Ctrl while clicking on the desired rows.
- To unhide multiple rows,then click on unhide in `format` section.
- Same for column hiding
- to select blank cells at once,select the region, `ctrl + g`,->`Special`->`Blanks`

### Paste's different features:

- **Normal pasting** : Paste everything (format,font,formulae)
- **Formula F pasting** : table data and corresponding formula
- **Formula and Number formatting pasting** : Table data ,formula and also number format (e.g, 56.00 -> 56.00)
- **Keep source formatting** : It pastes everything aside from few things such as width of the a specific column.
- **No Border Pasting**: As the name suggests,pasting will occur without any border.
- **Keep source column width**:
- **Transpose** :Column and row interchanges their place
- **Value and source Formatting** : Doesn't copy formula
- **Formatting paste**: doesn't paste data ,only format
- **Paste Link** : Change in pasted table will make change in sorce table.
- **Picture** : Pasted as Image format.
- **Picture Link format**: Same as `Picture` pasting .but update in main table will update the image as well.

### Protect the Excel :

- **Protecting sheet**:`Review `->`Protect`->`Protect sheet`->Enter password -> `ok`
- To make changes in protected sheet,You must `Unprotect` it.
- **Protecting Workbook**:`Review `->`Protect`->`Protect Workbook`->Enter password -> `ok`
- **Protecting a particular cell**:select the whole region,then `Home`->`Number` section drop down menu-> `Protection` -> uncheck the `lock`.then select the intended cell and ,then `Home`->`Number` section drop down menu-> `Protection` -> check the `lock`.After that `Review `->`Protect`->`Protect sheet`->Enter password -> `ok`
- **Lock during save**: During saving the file,click on `tools` (Left of the `Save` button).and then click on `general options`

#### Spell check:

- It can be found in `review` tab
- Adding comment also can be found in `review` section.

#### Array function:

```excel
= SUM(<Select a range> * <Select another range> )
<!-- -> `ctrl+shift+enter` -->
```

### AVERAGEIF, AVERAGEIFS and DAVERAGE

```excel
=AVERAGEIF(<range>,<criteria>,<Avg_Range>)
```

```excel
=AVERAGEIFS(<Avg_Range>,<range1>,<criteria1>,<range2>,<criteria2>....)
```

## Example and Explanation of the DAVERAGE Function in Excel

### Dataset:

| Name of Employee | Department   | Age | Sales |
| ---------------- | ------------ | --- | ----- |
| Murari Lal       | Blower       | 25  | 79    |
| Bagat Singh      | Mobile Phone | 28  | 67    |
| Raja             | LCD TV       | 30  | 47    |
| Utkal Kumar      | LCD TV       | 27  | 57    |
| Ajay Sharma      | LCD TV       | 26  | 48    |
| Ram Kumar        | Blower       | 25  | 94    |
| Vijay            | Blower       | 28  | 94    |
| Susmita          | LCD TV       | 25  | 49    |
| Bansal           | Mobile Phone | 29  | 76    |
| Sonu             | Mobile Phone | 26  | 47    |
| Kavita           | Mobile Phone | 29  | 58    |
| Rohan            | Mobile Phone | 27  | 97    |
| Mohan            | Blower       | 28  | 96    |
| Golu             | Mobile Phone | 24  | 48    |
| Avinita          | LCD TV       | 27  | 36    |
| Sakshi           | LCD TV       | 28  | 48    |
| Vimal            | Mobile Phone | 28  | 30    |
| Kamal            | Blower       | 29  | 63    |
| Roshan           | Blower       | 27  | 78    |

### Criteria Table (G1:H2):

| Department | Age |
| ---------- | --- |
| Blower     | >27 |

The criteria are:

- Employees who work in the "Blower" department.
- Employees whose age is greater than 27.

### Formula:

In cell G5, the following formula is entered:

```excel
=DAVERAGE(A1:D20, "Sales", G1:H2)
```

## DGET Formula:

```md
| S.no | Product Name     | Company    | Stock Units |
| ---- | ---------------- | ---------- | ----------- |
| 1    | Motor            | Crompton   | 56          |
| 2    | Heater           | Havells    | 75          |
| 3    | Fridge           | Godrej     | 34          |
| 4    | Cooler           | Bajaj      | 39          |
| 5    | Mixer Grinder    | Philips    | 89          |
| 6    | TV               | Samsung    | 70          |
| 7    | Remote           | xyz        | 65          |
| 8    | Speaker          | JBL        | 24          |
| 9    | Induction Chulla | Usha       | 20          |
| 10   | Microwave        | Haier      | 20          |
| 11   | AC               | LG         | 35          |
| 12   | Washing Machine  | IFB        | 67          |
| 13   | Exhaust Fan      | Electrolux | 68          |
| 14   | Ceiling Fan      | Khaitan    | 98          |
| 15   | Table Fan        | Orient     | 56          |
```

## Excel `DGET` Function Cheat Sheet

**Formula:**

```excel
=DGET(database, field, criteria)
```

### Explanation:

- **Database**: The entire range of the table including headers (e.g., `A1:D16`).
- **Field**: The column from which to retrieve data. This can either be the header name in quotation marks or the index of the column (e.g., `"Stock Units"` or `4`).
- **Criteria**: The range that specifies the conditions (e.g., `G1:G2`).

### Example:

```excel
=DGET(A1:D16, D1, G1:G2)
```

- **Database**: `A1:D16` - the range covering the entire table.
- **Field**: `D1` - refers to the "Stock Units" column.
- **Criteria**: `G1:G2` - matches the product name in the "Product Name" column.

### Key Notes:

- The `DGET` function retrieves a single value from a database that matches the criteria.
- If no matching record or more than one match is found, it returns an error.

## DMIN and DMAX

| S.no | Garment Type | Material | Size | Color | Total Piece |
| ---- | ------------ | -------- | ---- | ----- | ----------- |
| 1    | Kurta        | Cotton   | S    | Black | 20          |
| 2    | Pajama       | Rayon    | M    | White | 30          |
| 3    | Mens Jeans   | Cotton   | XL   | Blue  | 55          |
| 4    | Pants        | Cotton   | XXL  | White | 43          |
| 5    | Trouser      | Cotton   | L    | Green | 24          |
| 6    | Kurta        | Rayon    | XL   | Black | 65          |
| 7    | Pajama       | Cotton   | S    | White | 19          |
| 8    | Mens Jeans   | Rayon    | S    | Blue  | 8           |
| 9    | Pants        | Cotton   | M    | White | 39          |
| 10   | Trouser      | Rayon    | M    | Green | 26          |
| 11   | T Shirt      | Cotton   | N/A  | Pink  | 56          |
| 12   | T Shirt      | Silk     | N/A  | Grey  | 34          |
| 13   | Kurta        | Cotton   | XXL  | Black | 67          |
| 14   | Mens Jeans   | Cotton   | M    | White | 43          |
| 15   | Mens Jeans   | Cotton   | M    | Grey  | 9           |
| 16   | Mens Jeans   | Rayon    | M    | Grey  | 12          |
| 17   | Pajama       | Cotton   | XL   | Grey  | 46          |

---

## Excel `DMIN` and `DMAX` Cheat Sheet

### `DMAX` Function

**Formula:**

```excel
=DMAX(database, field, criteria)
```

- **Database**: The range of cells that contains the entire table.
- **Field**: The column to find the maximum value from. This can be the column name or the column index number.
- **Criteria**: The range where the conditions for the query are set.

### Example for Maximum in the Image:

In cell `I3`, the formula likely calculates the maximum value of **Total Piece** for the criteria provided:

```excel
=DMAX(A1:E17, E1, H1:I2)
```

- **Database**: `A1:E17` (the range of the table).
- **Field**: `E1` ("Total Piece" column).
- **Criteria**: `H1:I2` (criteria for finding the maximum value for Garment Type = "Kurta" and Material = "Cotton").

Result: 67

---

### `DMIN` Function

**Formula:**

```excel
=DMIN(database, field, criteria)
```

- **Database**: The range of cells containing the table.
- **Field**: The column to find the minimum value from. This can be the column name or the column index number.
- **Criteria**: The range of cells containing the criteria for the query.

### Example for Minimum in the Image:

In cell `I9`, the formula likely calculates the minimum value of **Total Piece** for the criteria provided:

```excel
=DMIN(A1:E17, E1, H8:I9)
```

- **Database**: `A1:E17` (range of the table).
- **Field**: `E1` ("Total Piece" column).
- **Criteria**: `H8:I9` (criteria for finding the minimum value for Garment Type = "Mens Jeans" and Size = "M").

Result: 9

---

### Key Notes for `DMAX` and `DMIN`:

- **DMAX** finds the maximum value in a field that meets the specified criteria.
- **DMIN** finds the minimum value in a field that meets the specified criteria.
- The **criteria** range is important and must match column headings and data conditions correctly.
  Here's a cheat sheet in Markdown format for the `TODAY`, `EDATE`, and `NOW` functions in Microsoft Excel:

````markdown
# Excel Functions Cheat Sheet

## 1. TODAY Function

**Description:**  
Returns the current date. The date updates automatically each time the worksheet is recalculated.

**Syntax:**

```excel
TODAY()
```
````

**Example:**

```excel
=TODAY()
```

_Returns today's date, e.g., `2024-10-11`._

---

## 2. EDATE Function

**Description:**  
Returns the serial number of the date that is the indicated number of months before or after a specified date.

**Syntax:**

```excel
EDATE(start_date, months)
```

- **start_date:** The starting date.
- **months:** The number of months to add (positive number) or subtract (negative number).

**Example:**

```excel
=EDATE(TODAY(), 3)
```

_Returns the date three months from today._

```excel
=EDATE(TODAY(), -1)
```

_Returns the date one month before today._

---

## 3. NOW Function

**Description:**  
Returns the current date and time. The value updates automatically each time the worksheet is recalculated.
**Example:**

```excel
=NOW()
```

_Returns the current date and time, e.g., `2024-10-11 14:30:00`._

---

Here’s a cheat sheet in Markdown format for the `COUNT`, `COUNTA`, `COUNTBLANK`, `COUNTIF`, `COUNTIFS`, and `DCOUNT` functions in Microsoft Excel:

## 1. COUNT Function

**Description:**  
Counts the number of cells that contain numeric values in a range.

**Syntax:**

```excel
COUNT(value1, [value2], ...)
```

- **value1:** The first argument that can be a range or a value.
- **value2:** (Optional) Additional arguments.

**Example:**

```excel
=COUNT(A1:A10)
```

_Counts the number of numeric values in the range A1:A10._

---

## 2. COUNTA Function

**Description:**  
Counts the number of non-empty cells in a range, regardless of the data type.

**Syntax:**

```excel
COUNTA(value1, [value2], ...)
```

- **value1:** The first argument that can be a range or a value.
- **value2:** (Optional) Additional arguments.

**Example:**

```excel
=COUNTA(A1:A10)
```

_Counts all non-empty cells in the range A1:A10._

---

## 3. COUNTBLANK Function

**Description:**  
Counts the number of empty cells in a specified range.

**Syntax:**

```excel
COUNTBLANK(range)
```

- **range:** The range in which to count empty cells.

**Example:**

```excel
=COUNTBLANK(A1:A10)
```

_Counts the number of empty cells in the range A1:A10._

---

## 4. COUNTIF Function

**Description:**  
Counts the number of cells that meet a single specified criterion.

**Syntax:**

```excel
COUNTIF(range, criteria)
```

- **range:** The range of cells to count.
- **criteria:** The condition that must be met.

**Example:**

```excel
=COUNTIF(A1:A10, ">10")
```

_Counts the number of cells in the range A1:A10 that are greater than 10._

---

## 5. COUNTIFS Function

**Description:**  
Counts the number of cells that meet multiple specified criteria.

**Syntax:**

```excel
COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)
```

- **criteria_range1:** The first range to evaluate.
- **criteria1:** The condition for the first range.
- **criteria_range2, criteria2:** (Optional) Additional ranges and conditions.

**Example:**

```excel
=COUNTIFS(A1:A10, ">10", B1:B10, "<5")
```

_Counts the number of rows where values in A1:A10 are greater than 10 and values in B1:B10 are less than 5._

---

## 6. DCOUNT Function

**Description:**  
Counts the cells that contain numbers in a database that meet specified conditions.

**Syntax:**

```excel
DCOUNT(database, field, criteria)
```

- **database:** The range that makes up the database, including headers.
- **field:** The column to count (can be the column label or the column number).
- **criteria:** The range that contains the criteria for counting.

**Example:**

```excel
=DCOUNT(A1:C10, "Sales", E1:E2)
```

_Counts the number of numeric entries in the "Sales" column of the range A1:C10 based on criteria specified in E1:E2._

---

### Usage Tips

- Use `COUNT` to get a quick count of numeric values.
- Use `COUNTA` when you need to include all types of entries, including text.
- Use `COUNTBLANK` to identify gaps in data entry.
- Use `COUNTIF` and `COUNTIFS` for conditional counting.
- Use `DCOUNT` for counting entries in a structured database format.

# Excel Functions Cheat Sheet

## 1. ROUND Function

**Description:**  
Rounds a number to a specified number of digits.

**Syntax:**

```excel
ROUND(number, num_digits)
```

- **number:** The number you want to round.
- **num_digits:** The number of digits to which you want to round the number.

**Example:**

```excel
=ROUND(3.14159, 2)
```

_Rounds 3.14159 to 2 decimal places, resulting in 3.14._

---

## 2. ROUNDDOWN Function

**Description:**  
Rounds a number down towards zero to a specified number of digits.

**Syntax:**

```excel
ROUNDDOWN(number, num_digits)
```

- **number:** The number you want to round down.
- **num_digits:** The number of digits to which you want to round down.

**Example:**

```excel
=ROUNDDOWN(3.14159, 2)
```

_Rounds 3.14159 down to 2 decimal places, resulting in 3.14._

---

## 3. FACT Function

**Description:**  
Calculates the factorial of a number.

**Syntax:**

```excel
FACT(number)
```

- **number:** The non-negative integer for which you want the factorial.

**Example:**

```excel
=FACT(5)
```

_Calculates the factorial of 5 (5!) which equals 120._

---

## 4. DISCOUNT Function

**Description:**  
Calculates the discount amount given a certain percentage off a price.

**Syntax:**

```excel
DISCOUNT(price, discount_rate)
```

- **price:** The original price of the item.
- **discount_rate:** The discount percentage expressed as a decimal.

**Example:**

```excel
=DISCOUNT(100, 0.2)
```

_Calculates a discount of 20% on $100, resulting in $20._

---

## 5. MOD Function

**Description:**  
Returns the remainder after a number is divided by a divisor.

**Syntax:**

```excel
MOD(number, divisor)
```

- **number:** The number you want to divide.
- **divisor:** The number by which you want to divide.

**Example:**

```excel
=MOD(10, 3)
```

_Returns the remainder of 10 divided by 3, which is 1._

---

## 6. EVEN Function

**Description:**  
Rounds a number up to the nearest even integer.

**Syntax:**

```excel
EVEN(number)
```

- **number:** The number you want to round up to the nearest even integer.

**Example:**

```excel
=ISEVEN(3) #returns true or false
```

_Rounds 3 up to the nearest even integer, resulting in 4._

---

## 7. ODD Function

**Description:**  
Rounds a number up to the nearest odd integer.

**Syntax:**

```excel
ODD(number)
```

- **number:** The number you want to round up to the nearest odd integer.

**Example:**

```excel
=ODD(4)
```

_Rounds 4 up to the nearest odd integer, resulting in 5._

---

### Usage Tips

- Use `ROUND` for general rounding needs.
- Use `ROUNDDOWN` when you want to ensure a number is rounded down.
- Use `FACT` for combinatorial calculations or probability problems.
- Use `DISCOUNT` to quickly calculate discount amounts for pricing.
- Use `MOD` for determining evenness or oddness of numbers or for cycling through values.
- Use `EVEN` and `ODD` to standardize data to the nearest whole number for certain calculations.

```

```
