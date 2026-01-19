/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Ensure Office is ready
Office.onReady(() => {
    document.getElementById("btn").onclick = run;
    
});

export async function run() {
  try {
        // Use a Callback function
        await Excel.run(async (context) => {
          // Get the currently active worksheet in the workbook: ws -> worksheet
          // const ws = context.workbook.worksheets.getActiveWorksheet();
          
          // Get the dashboard and the raw data sheets
          const rawDataSheet = context.workbook.worksheets.getItem("Raw Data");
          const dashboardSheet = context.workbook.worksheets.getItem("Dashboard");
            
          // STEP 1: Add a formula of Total Revenue to cell C8 in the Dashboard sheet
            // Define the range for the first step: add a formula to C8
            const cellC8 = dashboardSheet.getRange("C8");
            // Set the formula for the cell 8
            cellC8.formulas = [["=SUMIFS('Raw Data'!$E$2:$E$187,'Raw Data'!$D$2:$D$187,$A8,'Raw Data'!$B$2:$B$187,VALUE(LEFT($B8,4)),'Raw Data'!$C$2:$C$187,RIGHT($B8,2))"]];

             // Format as Currency with 0 decimals
            cellC8.numberFormat = [["$#,##0"]];
            await context.sync();

            // First check if C8 has the formula correctly
            //cellC8.load("values");
            // Copy and paste the formula from C8 to the range C9:C39
            const targetRangeCellC9toC39 = dashboardSheet.getRange("C9:C39");
            targetRangeCellC9toC39.copyFrom(cellC8, Excel.RangeCopyType.formulas);
            // Format as Currency with 0 decimals
            targetRangeCellC9toC39.numberFormat = [["$#,##0"]];
            await context.sync();



          // STEP 2: Weighted Average Margin (Column D)
            // Define the range for the second step: add a formula to D8
            const cellD8 = dashboardSheet.getRange("D8");
            // Set the formula for the cell D8
            cellD8.formulas = [["=SUMPRODUCT(('Raw Data'!$D$2:$D$187=$A8)*('Raw Data'!$B$2:$B$187=(LEFT($B8,4)))*('Raw Data'!$C$2:$C$187=RIGHT($B8,2))*'Raw Data'!$G$2:$G$187)/C8"]];
            // Format as Percentage with 1 decimals
            cellD8.numberFormat = [["0.0%"]];
            await context.sync();

            // Copy and paste the formula from D8 to the range D9:D39
            const targetRangeCellD9toD39 = dashboardSheet.getRange("D9:D39");
            targetRangeCellD9toD39.copyFrom(cellD8, Excel.RangeCopyType.formulas);
            // Format as Percentage with 1 decimals
            targetRangeCellD9toD39.numberFormat = [["0.0%"]];
            await context.sync();


          // STEP 3: Rolling 3-Month Trend (Column E)
            // Define the range for the third step: add a formula to E8
            const cellE8 = dashboardSheet.getRange("E8");
            // Set the formula for the cell E8
            cellE8.formulas = [["=IF(AND(LEFT($B8,4)=\"2023\",RIGHT($B8,2)=\"Q1\"),\"N/A\",D8-D7)"]];
            // Format as Percentage with 1 decimals
            cellE8.numberFormat = [["0.0%"]];
            await context.sync();

            // Copy and paste the formula from E8 to the range E9:E39
            const targetRangeCellE9toE39 = dashboardSheet.getRange("E9:E39");
            targetRangeCellE9toE39.copyFrom(cellE8, Excel.RangeCopyType.formulas);
            // Format as Percentage with 1 decimals
            targetRangeCellE9toE39.numberFormat = [["0.0%"]];
            await context.sync();

          // STEP 4: Year-over-Year Margin Delta (Column F)
            // Define the range for the third step: add a formula to F8
            const cellF8 = dashboardSheet.getRange("F8");
            // Set the formula for the cell F8
            cellF8.formulas = [["=IF(LEFT($B8,4)=\"2023\",\"N/A\",D8-INDEX($D$8:$D$39,MATCH($A8&LEFT($B8,4)-1&RIGHT($B8,2),$A$8:$A$39&LEFT($B$8:$B$39,4)&RIGHT($B$8:$B$39,2),0)))"]];
            // Format as Percentage with 1 decimals
            cellF8.numberFormat = [["0.0%"]];
            await context.sync();

            // Copy and paste the formula from F8 to the range F9:F39
            const targetRangeCellF9toF39 = dashboardSheet.getRange("F9:F39");
            targetRangeCellF9toF39.copyFrom(cellF8, Excel.RangeCopyType.formulas);
            // Format as Percentage with 1 decimals
            targetRangeCellF9toF39.numberFormat = [["0.0%"]];
            await context.sync();

          // STEP 5: Margin Health Classification (Column G)
            // Define the range for the third step: add a formula to G8
            const cellG8 = dashboardSheet.getRange("G8");
            // Set the formula for the cell F8
            cellG8.formulas = [["=IF(D8>0.35,\"Strong\",IF(D8>=0.2,\"Moderate\",\"At Risk\"))"]];
            // Conditional Formatting  
            await context.sync();

            // Copy and paste the formula from G8 to the range G9:G39
            const targetRangeCellG9toG39 = dashboardSheet.getRange("G9:G39");
            targetRangeCellG9toG39.copyFrom(cellG8, Excel.RangeCopyType.formulas);
            await context.sync();
            
            // Conditional Formatting
            // First remove all the conditional formatting from the used range for "Strong", "Moderate" and "At Risk"
            dashboardSheet.getUsedRange().conditionalFormats.clearAll();
            await context.sync();

            const strong = dashboardSheet.getRange("G8:G39").conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
            strong.cellValue.format.fill.color = "#C6EFCE";
            strong.cellValue.format.font.color = "#006100";
            strong.cellValue.rule = {
                formula1: "=\"Strong\"",
                operator: Excel.ConditionalCellValueOperator.equalTo
            };
            
            const moderate = dashboardSheet.getRange("G8:G39").conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
            moderate.cellValue.format.fill.color = "#FFEB9C";
            moderate.cellValue.format.font.color = "#9C5700";
            moderate.cellValue.rule = {
                formula1: "=\"Moderate\"",
                operator: Excel.ConditionalCellValueOperator.equalTo
            };  

            const atRisk = dashboardSheet.getRange("G8:G39").conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
            atRisk.cellValue.format.fill.color = "#FFC7CE";
            atRisk.cellValue.format.font.color = "#9C0006";
            atRisk.cellValue.rule = {
                formula1: "=\"At Risk\"",
                operator: Excel.ConditionalCellValueOperator.equalTo
            };  
            await context.sync();

          // STEP 6: Set Up Chart Data Table
            // Select the cellA42
            const cellA42 = dashboardSheet.getRange("A42");
            // Load the format property first
            cellA42.load("format");
            // Set the value for the cell A42
            cellA42.values = [["Chart Data"]];
            await context.sync();

            // Define the styling for the cell42
            cellA42.format.font.size = 11;
            cellA42.format.font.bold = true;
            cellA42.format.font.name = "Calibri";


            // Create the headers for the chart data table
            const headers = [["Quarter", "Widget Pro", "Widget Standard", "Service Package", "Accessory Kit", "Total Revenue"]];
            
            const chartHeaderRange = dashboardSheet.getRange("A43:F43"); 
            chartHeaderRange.values = headers;
            
            for(const header of headers){

                // Set the header values
                chartHeaderRange.values = headers;
                // Define the styling for the headers
                chartHeaderRange.format.font.size = 11;
                chartHeaderRange.format.font.bold = true;
                chartHeaderRange.format.font.name = "Calibri";
                chartHeaderRange.format.fill.color = "#D9E1F2";
            }
            

            await context.sync();

            // Create the quarters values for the chart data table
            const quarters = ["2023 Q1", "2023 Q2", "2023 Q3", "2023 Q4", "2024 Q1", "2024 Q2", "2024 Q3", "2024 Q4"];

            // Write each quarter individually if you need per-cell formatting
            for(let i = 0; i < quarters.length; i++) {
                const row = 44 + i; // Start at row 44
                const cell = dashboardSheet.getRange(`A${row}`);
                console.log(`A${row}`);
                cell.values = [[quarters[i]]];

            await context.sync();
        
          // STEP 7: Add Chart Data Formulas
            // Define the range for the third step: add a formula to B44
            const cellB44 = dashboardSheet.getRange("B44");
            // Set the formula for the cell B44
            cellB44.formulas = [["=SUMPRODUCT(($A$8:$A$39=\"Widget Pro\")*($B$8:$B$39=$A44)*($D$8:$D$39))"]];
            // Formatting
            cellB44.numberFormat = [["0.0%"]];  
            await context.sync();

            // Define the range for the third step: add a formula to C44
            const cellC44 = dashboardSheet.getRange("C44");
            // Set the formula for the cell C44
            cellC44.formulas = [["=SUMPRODUCT(($A$8:$A$39=\"Widget Standard\")*($B$8:$B$39=$A44)*($D$8:$D$39))"]];
            // Formatting
            cellC44.numberFormat = [["0.0%"]];  
            await context.sync();

            // Define the range for the third step: add a formula to D44
            const cellD44 = dashboardSheet.getRange("D44");
            // Set the formula for the cell D44
            cellD44.formulas = [["=SUMPRODUCT(($A$8:$A$39=\"Service Package\")*($B$8:$B$39=$A44)*($D$8:$D$39))"]];
            // Formatting
            cellD44.numberFormat = [["0.0%"]];  
            await context.sync();

            // Define the range for the third step: add a formula to E44
            const cellE44 = dashboardSheet.getRange("E44");
            // Set the formula for the cell E44
            cellE44.formulas = [["=SUMPRODUCT(($A$8:$A$39=\"Accessory Kit\")*($B$8:$B$39=$A44)*($D$8:$D$39))"]];
            // Formatting
            cellE44.numberFormat = [["0.0%"]];  
            await context.sync();

            // Define the range for the third step: add a formula to F44
            const cellF44 = dashboardSheet.getRange("F44");
            // Set the formula for the cell F44
            cellF44.formulas = [["=SUMIF($B$8:$B$39,$A44,$C$8:$C$39)"]];
            // Formatting
            cellF44.numberFormat = [["$#,##0"]];  
            await context.sync();
            
            // Copy and paste the formulas from row 44 to the range row 45 to row 51
            const targetRangeRow45to51 = dashboardSheet.getRange("B45:F51");
            targetRangeRow45to51.copyFrom(dashboardSheet.getRange("B44:F44"), Excel.RangeCopyType.formulas);

            // 1. Format B44:E51 as Percentage with 1 decimal
            const chartPercentageRange = dashboardSheet.getRange("B44:E51");
            chartPercentageRange.numberFormat = [["0.0%"]];
            
            // 2. Format F44:F51 as Currency with 0 decimals
            const chartCurrencyRange = dashboardSheet.getRange("F44:F51");
            chartCurrencyRange.numberFormat = [["$#,##0"]];
            }
          
          // STEP 8: Create the Chart
            // Select A43:F51 (headers + all data)
            const chartDataRange = dashboardSheet.getRange("A43:F51");
            // Create a clustered column chart
            const chart = dashboardSheet.charts.add(Excel.ChartType.columnClustered, chartDataRange, Excel.ChartSeriesBy.columns);
            // Set the position of the chart
            chart.setPosition("A53", "H53");
            // Add the chart title and set the height of the chart
            chart.title.text = "Quarterly Margin Trends by Product";
            chart.title.visible = true;
            chart.height = 300;

            // Load the chart series collection
            chart.series.load("items/name");

            await context.sync();
            
            // Variable pour les totalRevenues
            const seriesCollection = chart.series;
            // Initialize variable for Total Revenue series
            let totalRevenueSeries = null;

            for(let i = 0; i < seriesCollection.items.length; i++) {
                const series = seriesCollection.items[i];
                if(seriesCollection.items[i].name === "Total Revenue") {
                    // Format the Total Revenue series as a line chart
                    totalRevenueSeries = seriesCollection.items[i];
                    break;
                }
            }

            if(totalRevenueSeries) {
                // Change the axis group to secondary
                totalRevenueSeries.axisGroup = Excel.ChartAxisGroup.secondary;

                // Change the chart type to line
                totalRevenueSeries.chartType = Excel.ChartType.line;
                // Set the chart's line color and it's weight
                totalRevenueSeries.format.line.color = "red";
                totalRevenueSeries.format.line.weight = 2;
                console.log("Total Revenue series formatted as line chart.");

                await context.sync();
            }

            // Retrieves the primary axis and places it in a variable so that it can be modified and its properties managed: title, visibility, etc.
            chart.axes.valueAxis.title.text = "Profit Margin";
            chart.axes.valueAxis.visible = true;

            // Retrieves the secondary axis and places it in a variable so that it can be modified and its properties managed: title, visibility, etc.
            const secondaryAxis = chart.axes.getItem(Excel.ChartAxisType.value, Excel.ChartAxisGroup.secondary);
            secondaryAxis.title.text = "Total Revenue ($)";
            secondaryAxis.visible = true; 

            // Load position and width properties for columns A and F
            // This prepares the columns for positioning elements between them
            const colA = dashboardSheet.getRange("A:A"); // Get the entire column A
            const colF = dashboardSheet.getRange("F:F"); // Get the entire column F

            // Load the 'left' property (x-coordinate position) and 'columnWidth' 
            // These properties will be used to calculate positioning between columns
            colA.load("left, columnWidth"); // Load left position and width of column A
            colF.load("left, columnWidth"); // Load left position and width of column F

            await context.sync();

            // Adjust the position
            const leftPosition = colA.left;
            
            // Calculate the total width from column A to column F
            let totalWidth = 0;

            // Define the columns from A to F
            const cols = ["A", "B", "C", "D", "E", "F"];

            for(const col of cols) {
                // Get the range for each column
                const colRange =dashboardSheet.getRange(`${col}:${col}`);
                // Load the column width property
                colRange.format.load("columnWidth");
                await context.sync();
                
                // Add the column width to the total width
                const chartWidth = colRange.format.columnWidth > 0 ? colRange.format.columnWidth : 8.43; // Default width if not set: 8.43
                // Accumulate the total width
                totalWidth += chartWidth;
            }

            // Get the top position of cell A53
            const cellA53 = dashboardSheet.getRange("A53");
            // Load the top and rowHeight properties
            cellA53.load("top, rowHeight");

            await context.sync();

            // Get the top position
            const topPosition = cellA53.top;
            
            // Set the chart position and width
            chart.left = leftPosition;
            chart.top = topPosition;
            chart.width = totalWidth;

            await context.sync();
        });
    } catch (error) {
        console.error(error);
    };
}

