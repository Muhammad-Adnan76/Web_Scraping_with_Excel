# Web_Scraping_with_Excel
To scrape and visualize Paris Olympics medals data using Excel's Power Query, follow these detailed steps:

### **1. Find a Data Source**

Identify a reliable website that provides detailed medal data for the Paris Olympics. For this example, let’s assume you have a URL like `https://www.example.com/paris-olympics-medals`. Ensure the data is in a structured format such as HTML tables.

### **2. Import Data Using Power Query**

**Step 1: Open Excel**

- Launch Excel and open a new workbook or an existing one where you want to import the data.

**Step 2: Access Power Query**

- Navigate to the `Data` tab on the Ribbon.
- Select `Get Data` > `From Other Sources` > `From Web`.

**Step 3: Enter URL**

- In the `From Web` dialog box, enter the URL of the web page with the medals data (e.g., `https://www.example.com/paris-olympics-medals`).
- Click `OK`.

**Step 4: Load Data**

- Excel will connect to the web page and load a preview of the available data.
- If your data is in a table format, you will see a list of tables in the `Navigator` window.
- Select the table with the medals data and click `Load` to import it directly into your workbook.

  If you need to transform the data (e.g., remove unnecessary columns, filter data), click `Transform Data`. This opens the Power Query Editor.

**Step 5: Transform Data (Optional)**

- In the Power Query Editor, you can clean and transform the data as needed:
  - **Remove Unnecessary Columns:** Right-click on column headers and choose `Remove`.
  - **Filter Data:** Click on column drop-downs to filter data based on criteria.
  - **Rename Columns:** Double-click on column headers to rename them.

- After making the necessary transformations, click `Close & Load` to load the data into your Excel worksheet.

### **3. Visualize the Data**

**Step 1: Create a Pivot Table**

- With your data loaded into Excel, go to the `Insert` tab on the Ribbon.
- Select `PivotTable` from the options. Excel will prompt you to choose the data range (the table you imported).
- Click `OK` to create the PivotTable in a new worksheet or an existing one.

**Step 2: Configure the Pivot Table**

- Drag fields from the `PivotTable Fields` pane to the `Rows`, `Columns`, and `Values` areas to organize your data.
  - **Rows:** Drag `Country` or `Medal Type` to the Rows area to categorize the data.
  - **Columns:** Drag `Year` or `Sport` to the Columns area if you want to compare medals across different categories.
  - **Values:** Drag `Medal Count` or equivalent to the Values area to aggregate the data.

**Step 3: Create Charts**

- With the PivotTable ready, you can now create charts to visualize the data.
- Click anywhere within the PivotTable.
- Go to the `Insert` tab and select the type of chart you want (e.g., Column Chart, Pie Chart, Bar Chart).

**Step 4: Customize Charts**

- After inserting a chart, you can customize it using the `Chart Tools` in the Ribbon.
  - **Add Titles:** Click on the chart title to add a descriptive title.
  - **Adjust Colors:** Use the `Format` tab to change colors and styles.
  - **Add Labels:** Add data labels or legends to make the chart more informative.

**Step 5: Format Data**

- For additional formatting, use Excel’s built-in formatting options:
  - **Conditional Formatting:** Highlight top-performing countries or medal types.
  - **Cell Styles:** Apply different cell styles to make your table more readable.

### **Example Workflow**

1. **Data Import:**
   - URL: `https://www.example.com/paris-olympics-medals`
   - Import table with medal data.

2. **Data Transformation:**
   - Clean data by removing unnecessary columns.
   - Filter to include only the current Olympics data.

3. **Pivot Table Setup:**
   - Rows: `Country`
   - Columns: `Medal Type`
   - Values: `Medal Count`

4. **Visualization:**
   - Create a bar chart to show the total number of medals by country.
   - Create a pie chart to visualize the distribution of medal types.

### **Summary**

Using Excel's Power Query and PivotTable features allows you to efficiently import, transform, and visualize data from the web. This approach is useful for analyzing and presenting Olympic medal data or similar datasets. If you encounter issues with the specific website structure, consider using more advanced data cleaning techniques or scripts.
