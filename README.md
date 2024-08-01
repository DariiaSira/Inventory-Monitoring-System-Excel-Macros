# Warehouse inventory monitoring system using macros in Excel

## Description
The project is the development of Excel macros for efficient inventory management in the warehouse. The system automates the process of inventory tracking, provides notifications on the need for replenishment, and simplifies the creation of purchase orders. The main goal of the project is to improve inventory management and reduce the risks associated with shortages.

## Project Objectives
1. **Inventory Update**: Automatically update current inventory based on daily receipts and shipments, ensuring stock information is up to date.
2. **Low Stock Notification:** Automatically notify the user when stock levels reach or fall below minimum levels, allowing for timely response to potential shortages.
3. **Automatic Purchase Order Creation:** Generate purchase orders for items below a set minimum quantity, simplifying the replenishment process.
4. **Report Generation:** Generate inventory movement reports, providing detailed information on receipts, shipments and current inventory status for analysis and planning.

## Key Functions
1. Automatic change of stock quantities based on receipts and shipments.
2. User notification when the minimum stock level is reached.
3. Automation of the process of ordering goods for stock replenishment.
4. Generating reports on inventory movement for analysis and accounting.

## Description of Worksheets

### Table 1: Stocks

**Purpose:** Storing current stock data in the warehouse.

- **Commodity Code**: A unique identifier for each item.
- **Name**: The name of the item.
- **Stock Quantity**: The current quantity of the item available in stock.
- **Minimum**: The minimum allowable quantity of an item at which replenishment is required.
- **Price per unit**: The cost per unit of an item.
- **Supplier**: The name of the company or person supplying the item.

**Example:**
| Item Code | Item Name | Quantity in Stock | Minimum Level | Unit Price | Supplier       |
|-----------|-----------|--------------------|---------------|------------|----------------|
| 1001      | Widget A  | 75                 | 20            | 12.5       | Alpha Supplies |
| 1002      | Widget B  | 35                 | 25            | 15         | Beta Trading   |
| 1003      | Gadget C  | 20                 | 15            | 25         | Gamma Goods    |

### Table 2: Receipts

**Purpose:** Record all receipts of goods into the warehouse.

- **Product Code**: The unique identifier of the item received into the warehouse.
- **Date**: The date the item was received.
- **Quantity**: The quantity of the item received.
- **Source**: The name of the company or person who made the delivery.

**Example:**
| Item Code | Date       | Quantity | Source          |
|-----------|------------|----------|-----------------|
| 1001      | 01/07/2024 | 50       | Alpha Supplies  |
| 1002      | 02/07/2024 | 20       | Beta Trading    |
| 1003      | 03/07/2024 | 30       | Gamma Goods     |

### Table 3: Shipments

**Purpose:** Record all shipments of goods from the warehouse.

- **Product Code**: Unique identifier of the item shipped from the warehouse.
- **Date**: The date the item was shipped.
- **Quantity**: The quantity of the item shipped.
- **Customer**: The name of the customer to whom the item was shipped.

**Example:**
| Item Code | Date       | Quantity | Customer       |
|-----------|------------|----------|----------------|
| 1001      | 01/07/2024 | 20       | Alpha Supplies |
| 1002      | 02/07/2024 | 10       | Beta Trading   |
| 1003      | 03/07/2024 | 25       | Gamma Goods    |

## Macros
### Macro for stock update
**Description:**
This macro handles incoming and outgoing shipments of goods, automatically updating the quantity in stock. It takes data from the Incoming Shipments and Outgoing Shipments tables, and adjusts the stock quantity in the Inventory table accordingly.

**Key Steps:**
1. Adds the quantity of incoming goods to the current quantity in the Inventory table.
2. Subtracts the quantity shipped from the current quantity in the Inventory table. If theinventory level is below the minimum level, a warning message is displayed.
3. If an item code is not found in the Inventory table, an error message is displayed.

### Macro for low stock notification
**Description:**
This macro checks stock levels in the warehouse and creates or updates a Stock Alerts sheet that displays items with stock levels below the minimum.

**Key Steps:**
1. Creates a new Stock Alerts sheet or clears an existing one to record alerts.
2. Compares the current stock level to the minimum level for each item in the Inventory table.
3. If the stock level is below the minimum level, adds the item information to the Stock Alerts sheet labeled “Needs Restocking”.

### Macro for creating purchase orders
**Description:**
This macro automatically creates purchase orders for items that need replenishment. It analyzes data from the Inventory table, determines which items have inventory levels below the minimum, and generates a list of needed orders in the Purchase Orders table.

**Key Steps:**
1. Clears the previous data on the Purchase Orders sheet.
2. Calculates the quantity that must be ordered to reach the minimum level for each item.
3. Records data about items requiring replenishment in the Purchase Orders sheet.

### Macro for generating reports:
**Description:**
This macro generates a report on the current status of inventory in the warehouse. It takes data from the Inventory table and generates a report in the Report table that includes information about item codes, item names, current quantities, and minimum levels.

**Key steps:**
1. Clears the previous data on the Report sheet.
2. Copies data from the Inventory table to the Report table, including item code, item name, current quantity, and minimum inventory levels.

### Macro for table formatting:
**Description:**
This macro formats a table on the active sheet by applying borders and alignment. It sets borders for all cells in the table and makes the first row with headings bold and center-aligned.

**Key Steps:**
1. Finds the table range on the active worksheet.
2. Sets borders for all sides of the table and internal cell borders.
3. Makes the first row bold and centers it horizontally and vertically.
4. Aligns all cells in the table horizontally and vertically in the center.

## Screenshots and Examples
There are screenshot of Inventory table before and after the stock update. We see that every item was changes regarding to information from Incoming and Outgoing Shipments tables.

![image](https://github.com/user-attachments/assets/908bb312-270e-439a-b29d-4d121520eb78)
![image](https://github.com/user-attachments/assets/3b057563-a8af-4ee9-9130-f8457a12cc8c)

There is screenshot of Macro for low stock notification performing. So, the new sheet with a table of needed items was created based on formated Inventory table. Also the Macro for formating was run.

![image](https://github.com/user-attachments/assets/ec227dab-84a5-4d69-8774-fe8f7de980e1)

Based the Inventory table the Macro for creating purchase orders calculated the quantity that must be ordered to reach the minimum level for each item. Also formated using Macro for formating.

![image](https://github.com/user-attachments/assets/136b4a6f-74e2-4f25-aff3-cc3efdedc89a)

The example of Report table using Macro for generating reports was generated.

![image](https://github.com/user-attachments/assets/c014e540-171f-409d-a8b4-87b943bb987b)

## Conclusion
The warehouse inventory monitoring project is a powerful and flexible inventory management solution that can significantly improve the efficiency of warehouse logistics. As part of this project, macros have been developed that provide automatic updates of current stock data, notify the user of the need to replenish stock and create purchase orders.

The main goal of the project is to create a user-friendly tool for tracking stock in the warehouse, minimizing the risk of stock-outs and ensuring timely replenishment. This is achieved by automating key inventory management processes such as updating data based on receipts and shipments, notifying the user when minimum stock levels are reached and automatically creating purchase orders. As a result, warehouse logistics becomes more predictable and manageable.

