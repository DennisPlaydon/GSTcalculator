# GSTcalculator

This is a custom solution built for a sole sole trader that has no accountant. 

Filing your own GST can be hard because you need to categorise your purchases. Doing so by hand can be quite tiresome and time intensive. 

## Solution

GST Calculator taxes a lot of the hard work out of filing for GST. 
It is a simple program that will output an excel sheet with highlighted columns to showcase payments and receipts. Running this through lives clients showed an average time saving of **40 minutes**

## How it works

- The user downloads their transactions to an excel sheet
- The program reads each line and tries to apply the transaction to a regular expression (a common spending category) and will the categorise the transaction
- The program will insert columns for money coming in and money leaving and highlight the cells it is not sure about.

## What I learnt

- Reading input from excel sheets
- Modifying columns and cells via python
- Regular expression matching
