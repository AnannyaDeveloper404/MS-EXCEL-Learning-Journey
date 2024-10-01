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

## Home Tab

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
