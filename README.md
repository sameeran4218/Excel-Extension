# 📊 Regression Analysis Macro for Excel

## 📄 Description

Easily perform regression analysis directly within Microsoft Excel with this VBA-powered tool. It automates the computation of slope, intercept, and R² score, handles missing data by replacing it with column medians, and generates insightful graphs to visualize results.

## Key Highlights:

- Automatically calculates regression metrics.

- Generates scatter plot and best fit regression line.

- Deployable as an Excel Add-In for seamless use.

## 💻 Tech Stack

Language: VBA (Visual Basic for Applications)

Platform: Microsoft Excel

File Types: .xlam (Excel Add-In)

## ✨ Features

📐 Automated Regression Analysis: Computes slope, intercept, and R² score effortlessly.

🛠️ Error Handling: Handles missing and non-numeric data.

📊 Graphical Output: Plots actual vs predicted values with a regression line.

🔄 Customizable: Specify input columns for independent and dependent variables.

🚀 Excel Add-In Support: Easily accessible across multiple workbooks.

## 📥 Installation Instructions

### Use as Add-In (.xlam)
- Download the .xlam file provided in the repository.
- Locate the file and go to Properties and check the Unblock option at the bottom of the window.
- Open Excel.
- Go to File > Options > Add-Ins.
- At the bottom, in the Manage dropdown, select Excel Add-ins and click Go.
- Click Browse and locate your .xlam file.
- Select the file and click OK. The add-in will now appear in the list of available add-ins.
- Go to Developer Tab > Excel Add-Ins and check the downloaded .xlam file.

## 📖 User Instructions

Run the Macro:
- Go to My Tab and click on Linear Regression Add-In
-Enter column letters for the independent variable (X) and dependent variable (Y) when prompted.
-Results are displayed next to the data.
-Predicted values appear in the subsequent column.
-A chart with actual values and the regression line is generated automatically.

## ⚠️ Error Handling

Missing Values: Automatically replaced with column medians.

Non-Numeric Data: Skipped with a warning.

Invalid Input: Prompts for valid column letters.

