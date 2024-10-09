# Instructions for Configuring the Visualization Script

This script is designed to generate a customized view table in Excel by joining data from a fact table and one or more dimension tables. It allows you to define the output columns, including calculated fields, and add slicers for interactive data exploration.

Follow the instructions below to configure the script according to your data and requirements.

---

## Overview

The main configuration object is `visualizationConfig`, which you need to customize:

```typescript
const visualizationConfig: VisualizationConfig = {
  outputSheetName: "Your Output Sheet Name",
  factTable: { /* Fact Table Configuration */ },
  dimensionTables: [ /* Dimension Tables Configuration */ ],
  outputColumns: [ /* Output Columns Configuration */ ],
  slicers: [ /* Slicers Configuration */ ],
};
```

Each section of the configuration object is explained in detail below.

---

## Configuration Sections

### 1. Output Sheet Name

**Purpose**: Specifies the name of the worksheet where the output data will be displayed.

- **Property**: `outputSheetName`
- **Example**:

  ```typescript
  outputSheetName: "Project Display",
  ```

---

### 2. Fact Table Configuration

**Purpose**: Defines the main data table containing the key metrics or facts.

- **Property**: `factTable`
- **Properties**:
  - `sheetName`: The name of the worksheet containing the fact table.
  - `keyColumns`: An array of column names to include from the fact table.

- **Example**:

  ```typescript
  factTable: {
    sheetName: "PROJECT_ITEMS",
    keyColumns: ["Line Item ID", "Quantity", "Project ID"],
  },
  ```

---

### 3. Dimension Tables Configuration

**Purpose**: Specifies additional tables that provide context or attributes (dimensions) to the fact data.

- **Property**: `dimensionTables`
- **Each Dimension Table Configuration Includes**:
  - `sheetName`: The name of the worksheet containing the dimension table.
  - `joinColumnFact`: The column name in the fact table to join on.
  - `joinColumnDim`: The column name in the dimension table to join on.
  - `selectColumns`: An array of column names to include from the dimension table.

- **Example**:

  ```typescript
  dimensionTables: [
    {
      sheetName: "PROJECTS",
      joinColumnFact: "Project ID",
      joinColumnDim: "Project ID",
      selectColumns: ["Project ID", "Project Name"],
    },
    {
      sheetName: "LINE_ITEMS",
      joinColumnFact: "Line Item ID",
      joinColumnDim: "Line Item ID",
      selectColumns: ["Line Item ID", "Line Item", "Description", "Price Act"],
    },
  ],
  ```

---

### 4. Output Columns Configuration

**Purpose**: Defines which columns will appear in the output sheet and how they are derived.

- **Property**: `outputColumns`
- **Each Output Column Configuration Includes**:
  - `header`: The column header in the output sheet.
  - `source`: Where the data comes from. Options:
    - `'fact'`: From the fact table.
    - `'dimension'`: From a dimension table.
    - `'calculated'`: Computed using a formula.
  - `sheetName`: (Required for `'dimension'`) The name of the dimension sheet.
  - `columnName`: The name of the column in the source sheet.
  - `formula`: (Required for `'calculated'`) A function to compute the value.
  - `type`: The data type for formatting. Options: `'STRING'`, `'NUMBER'`, `'CURRENCY'`, `'DATE'`.

- **Example**:

  ```typescript
  outputColumns: [
    {
      header: "Project Name",
      source: 'dimension',
      sheetName: "PROJECTS",
      columnName: "Project Name",
      type: "STRING",
    },
    {
      header: "Quantity",
      source: 'fact',
      sheetName: "PROJECT_ITEMS",
      columnName: "Quantity",
      type: "NUMBER",
    },
    {
      header: "Total",
      source: 'calculated',
      formula: (row) => {
        const quantity = Number(row["Quantity"]);
        const price = Number(row["Price Act"]);
        return quantity * price;
      },
      type: "CURRENCY",
    },
  ],
  ```

---

### 5. Slicers Configuration

**Purpose**: Adds interactive filters (slicers) to the output sheet for specified columns.

- **Property**: `slicers`
- **Each Slicer Configuration Includes**:
  - `columnName`: The name of the column to which the slicer will apply.

- **Example**:

  ```typescript
  slicers: [
    { columnName: "Project Name" },
    { columnName: "Line Item" },
  ],
  ```

---

## Step-by-Step Configuration Guide

### Step 1: Set the Output Sheet Name

Decide on a name for the output sheet and set the `outputSheetName` property.

```typescript
outputSheetName: "Your Output Sheet Name",
```

*Ensure this name does not conflict with existing sheet names unless you intend to overwrite.*

---

### Step 2: Configure the Fact Table

Identify your main data table and specify:

- **`sheetName`**: The worksheet name containing the fact table.
- **`keyColumns`**: The columns from the fact table that you want to include.

```typescript
factTable: {
  sheetName: "Your Fact Table Sheet Name",
  keyColumns: ["Column1", "Column2", "Column3"],
},
```

*Ensure all specified columns exist in the fact table sheet.*

---

### Step 3: Configure Dimension Tables

For each dimension table:

1. **Identify the table** and determine how it joins to the fact table.
2. **Specify**:
   - **`sheetName`**: The worksheet name containing the dimension table.
   - **`joinColumnFact`**: The column in the fact table used for joining.
   - **`joinColumnDim`**: The column in the dimension table used for joining.
   - **`selectColumns`**: The columns from the dimension table to include.

```typescript
dimensionTables: [
  {
    sheetName: "Dimension Sheet Name",
    joinColumnFact: "Fact Table Join Column",
    joinColumnDim: "Dimension Table Join Column",
    selectColumns: ["DimColumn1", "DimColumn2"],
  },
  // Add more dimension tables as needed
],
```

*Ensure the join columns exist in both the fact and dimension tables.*

---

### Step 4: Define Output Columns

Specify the columns to appear in the output sheet.

For each output column:

- **`header`**: The column header in the output sheet.
- **`source`**: `'fact'`, `'dimension'`, or `'calculated'`.
- **If `source` is `'fact'` or `'dimension'`**:
  - **`sheetName`** (for `'dimension'`): The sheet name.
  - **`columnName`**: The column name in the source sheet.
- **If `source` is `'calculated'`**:
  - **`formula`**: A function that calculates the value.
- **`type`**: The data type (`'STRING'`, `'NUMBER'`, `'CURRENCY'`, `'DATE'`).

```typescript
outputColumns: [
  // For columns from the dimension table
  {
    header: "Column Header",
    source: 'dimension',
    sheetName: "Dimension Sheet Name",
    columnName: "Dimension Column Name",
    type: "STRING",
  },
  // For columns from the fact table
  {
    header: "Column Header",
    source: 'fact',
    columnName: "Fact Column Name",
    type: "NUMBER",
  },
  // For calculated columns
  {
    header: "Calculated Column",
    source: 'calculated',
    formula: (row) => {
      // Your calculation logic
    },
    type: "CURRENCY",
  },
],
```

*Ensure that all referenced columns exist and are correctly named.*

---

### Step 5: Add Slicers (Optional)

To add interactive filters:

- Specify the columns you want to use as slicers.

```typescript
slicers: [
  { columnName: "Column Name for Slicer" },
  // Add more slicers as needed
],
```

*Slicers are optional and can enhance data exploration.*

---

## Tips and Best Practices

- **Column Names**: Match column names exactly as they appear in your worksheets, including case and spaces.
- **Data Types**:
  - Use `'STRING'` for text.
  - Use `'NUMBER'` for integers.
  - Use `'CURRENCY'` for monetary values (formatted with two decimal places).
  - Use `'DATE'` for date values (formatted as `mm/dd/yyyy`).
- **Calculated Columns**:
  - The `formula` function receives an object `row`, containing all joined data.
  - Handle any potential `undefined` or `null` values to prevent errors.
- **Error Handling**:
  - The script will throw an error if a specified column is not found. Double-check all names.
- **Testing**:
  - Start with a simple configuration and gradually add complexity.
  - Test the script after each change to ensure it works as expected.

---

## Full Example Configuration

```typescript
const visualizationConfig: VisualizationConfig = {
  outputSheetName: "Project Display",
  factTable: {
    sheetName: "PROJECT_ITEMS",
    keyColumns: ["Line Item ID", "Quantity", "Project ID"],
  },
  dimensionTables: [
    {
      sheetName: "PROJECTS",
      joinColumnFact: "Project ID",
      joinColumnDim: "Project ID",
      selectColumns: ["Project ID", "Project Name"],
    },
    {
      sheetName: "LINE_ITEMS",
      joinColumnFact: "Line Item ID",
      joinColumnDim: "Line Item ID",
      selectColumns: ["Line Item ID", "Line Item", "Description", "Price Act"],
    },
  ],
  outputColumns: [
    {
      header: "Project Name",
      source: 'dimension',
      sheetName: "PROJECTS",
      columnName: "Project Name",
      type: "STRING",
    },
    {
      header: "Item ID",
      source: 'dimension',
      sheetName: "LINE_ITEMS",
      columnName: "Line Item ID",
      type: "STRING",
    },
    {
      header: "Line Item",
      source: 'dimension',
      sheetName: "LINE_ITEMS",
      columnName: "Line Item",
      type: "STRING",
    },
    {
      header: "Description",
      source: 'dimension',
      sheetName: "LINE_ITEMS",
      columnName: "Description",
      type: "STRING",
    },
    {
      header: "Quantity",
      source: 'fact',
      columnName: "Quantity",
      type: "NUMBER",
    },
    {
      header: "Price",
      source: 'dimension',
      sheetName: "LINE_ITEMS",
      columnName: "Price Act",
      type: "CURRENCY",
    },
    {
      header: "Total",
      source: 'calculated',
      formula: (row) => {
        const quantity = Number(row["Quantity"]);
        const price = Number(row["Price Act"]);
        return quantity * price;
      },
      type: "CURRENCY",
    },
  ],
  slicers: [
    { columnName: "Project Name" },
    { columnName: "Line Item" },
  ],
};
```

---

## Running the Script

1. **Insert Your Configuration**: Replace the default `visualizationConfig` in the script with your customized configuration.
2. **Test the Script**: Run the script in Excel to generate the view table.
3. **Verify the Output**: Check the output sheet for accuracy.
4. **Adjust as Needed**: If you encounter errors or unexpected results, revisit your configuration.

---
