## Description

This is Joan Hsiao's portfolio.
I have 3+ years experience in UiPath development for finance and accounting, and implemented UiPath and Excel VBA training programs to the company.

## Table of Contents

* [RPA_UiPath](#rpa_uipath)
  * [Convert PDF Invoice to Excel Table](#convert-pdf-invoice-to-excel-table)
  * [Oracle ERP GL Account Entry](#oracle-erp-gl-account-entry)
  * [Maintenance E-billing in customer's website](#maintenance-e-billing-in-customers-website)
  * [DC Notes Entry](#dc-notes-entry)
* [VBA for Excel](#vba-for-excel)

# RPA_UiPath

# Convert PDF Invoice to Excel Table

* Input 
  * Any PDF file from ERP system, not including scanned copy.
  * Save input folder to same folder as robot.
* Process
  * Combine mutiple PDF files to one file.
  * Save as text file.
  * Parse text file by marco
  * Transfer to data table in excel.
* Maintenance
  * Edit marco by different PDF source (Ex. By vendor) 

# Oracle ERP GL Account Entry

* Input 
  * According to the sample, put the journal entry in file and save in "1_Excel" folder.
  * Robot will transfer one file to one journal entry.
* Process
  * Read all files "1_Excel" folder.
  * Open Oracle ERP.
  * Key-in journal entry one-by-one.
  * Close ERP system.

# Maintenance e-billing in customer's website

* Input 
  * According to the sample, download report and paste on worksheet.
* Process
  * Sort report by DO. number
  * Read data table.
  * Open customer's website and e-billing list page.
  * Open DO. one by one, check data, if correct then key-in information in field.
  * Confirm and save at website.
  * Write the status in excel file.
  
# DC Notes Entry
* Input 
  * Setting mail folder and OU(Operation Unit) in "Credit_Memo_File_Path.xlsm".
* Process
  * Download report from Oracle dicoverer system.
  * Get outlook mail and its attachments(credit/debit memo) from vendors.
  * Parse credit/debit memo to data table in excel by marco.
  * Mapping dicoverer report and credit/debit memo by number, then check date and amount, finally, write mapping status.
  * Generate a list of mapping correct.
  * Open Oracle ERP.
  * Find credit/debit memo by number one by one, and confirm them.
  * Write the entry status in excel file.
  * Generate status report and mail to manager and colleague.
  * Archive files by vendor.
* Maintenance
  * When add new vendor, add folder named by vendor code and add excel file with marco.
  * Edit marco by vendor.
  
# VBA for Excel
* FinancialReportDownload.xlsm
  * It's a web crawler write by VBA, for downloading financial report and market price of listed company from TWSE(Taiwan Stock Exchange Corporation) and TPEX(Taipei Exchange) website.