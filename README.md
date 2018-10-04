# vbsclasslists

This repository tracks the automated class lists project for Vacation Bible School at Mount Lebanon Evangelical Presbytarian Church.

This project was my first Python project. I learned everything on my own through Internet research. The purpose of this project was to help the Children's Ministry Director create class lists for VBS. During VBS, new class lists need to be created a few times before VBS starts as well as daily for the first two or three days of VBS as children are moved around to accommodate class size requirements and buddy requests. 

### createchildinfo.py and childinfofunctions.py
The code in createchildinfo.py reads the raw child information data file (VBSChildinformation2018_Data.xlsx), copies relevant columns, and writes the data to a master tab of a new file named VBSChildInformation2018.xlsx. This code also creates a new column 'Crew' for the director to assign each child to a crew. The childinfofunctions.py file contains code for formatting VBSChildInformation2018.xlsx. 

### createclasslists.py
The code in this file adds a tab for each crew to the VBSChildInformation2018.xlsx. On each tab, ths code writes the names of children assigned to the respective crew. If the tabs already exist, this code deletes children's names and writes new names to account for changes in crew assignment. This code does not overwrite the header rows on the tabs. The headers have hand-entered teacher and room information that does not typically change.

### createpaymentsum.py and paymentfunctions.py
The code in createpaymentsum.py reads the raw payment information data file (MLEPCVBSPayments2018_Data.xlsx), copies relevant columns, and writes the data along with payment balances to a new file MLEPCVBSPaymentBalances2018.xlsx. The paymentfunctions.py file contains code for formatting MLEP.CVBSPaymentBalances2018.xlsx.
