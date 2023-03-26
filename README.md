# Data-Warehousing-for-Insurance-Company-using-Excel-Microsoft
This repository contains the files and documentation for data warehousing project created for an insurance company using Excel Microsoft. 
Data Warehousing is a central repository of information that can be analyzed to make more informed decisions. For this project, we will create a simple Data Warehouse for an insurance company using Microsoft Excel. The finance department at the insurance company ran the Encounters for the year 2022 using the company’s server, and each month was saved as an individual file. The fiscal year started from July-Jun, in this case, July was month 1 of 2022. The Finance team forwarded 12 files for the year 2022 to the reporting group in order to create the Data Warehouse for the company starting from 2022.

To complete this task, we will follow the following steps:

## Step 1:

We will combine all 12 files in one Excel tab and name it "Raw Data".
We will add a column called "Encounter Month" and use a formula to calculate the month number of the encounter.
We will remove the test encounters that have the name "Xxtest" in the Sex column.
We will clean the Raw Data tab from any noisy data, using filters to find and delete encounters with noisy data.
We will remove any encounter that has a blank value in any column.
We will create a column "Discount's Region" to check which encounter has to get a discount. Encounters from southeast or southwest will get a discount. We will put 1 for encounters in these regions and 0 for others.
We will copy the Raw Data tab to two other tabs and name them "Company's Users" and "Billed Encounters".

## Step 2:
We will clean all three tabs by removing any noisy data and encounters with blank values.
We will keep the record with Insurance_ID 939361 with Encounter Date 10/14/2022.
We will unduplicate the records in the Company's Users tab using Insurance_ID as a key and keep the one that has the most recent Enrolment Date.
We will create a pivot table using the Company's Users tab to show the number of encounters for each month by region with the grand total.
We will create a pivot table using the Billed Encounters tab to show the number of billed services and enabling services encounters for each region with the grand total.
We will create another pivot table using the Billed Encounters tab to show how many region's encounters have to get a discount according to the company's decision and the total expenses for each region in the month and grand total. The total expenses will exclude the enabling services encounters.

## Step 3:
We will create a Summary table that gives all the answers of the above questions in an appropriate way.
Step 4:
We will choose a nice name for our report and write "FY 2022" in the second row, first cell.
We will create a Summary tab that helps our manager to read the report.
Step 5:
We will save all tabs and pivot tables and submit our work as an Excel file format.

#Link for the final dataware house in exel
https://docs.google.com/spreadsheets/d/13IarGZZ4aupo62IEZTii6B1oelYuGB69/edit?usp=share_link&ouid=117570485710671458515&rtpof=true&sd=true
