The goal is to produce a CSV file containing all inputs required to attach a router to a vManage device template without the user
needing to copy/paste and manually manipulate data

The script provides a CSV file that covers multiple templates as vManage will only pull out the required info from the CSV file

This python script takes an .xlxs workbook as input and reads and manipulates the data before outputting the results to a .csv file
The .csv contains all the variables required to attach routers to various templates within vManage

The script is customer specific as each customer will have different template variables
The idea is that this script provides a solid base and can be easily adapted on a customer basis

The next phase of development will utilise the SDWAN API to provide direct integration with vManage
