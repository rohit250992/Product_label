# Product_label
# Read and store content of an excel file.
 This part of the script reads your excel file and store all its contents. The path must be provided by the user to read the specific file.

# Write the dataframe object into csv file.
 Now to read the values of the file correctly and not just the formulas present in the cell. I have converted it into a CSV and again saved    it as an excel, so I can get the desired result. 

# Load the entire workbook & # Load one worksheet.
 This section of the script open ups the whole excel file as a workbook and BPR as a worksheet. Now, I saved all the values required to make the shipping label like Name, Daily Dose, capsule per bottle, lot number, manufacturing date and customer id for the worksheet. 

## Making supplement chart 
Now to make an example like supplement chart I have just used the Item name, percent and dosage from the formula sheet and put all the values in a table form. 

![table_image](https://github.com/rohit250992/Product_label/assets/117368239/13016f2f-6bea-418c-8961-341837f5836e)

#Creating a Template.

This section of the code is creating a template of the suggested shipping label. It inserts all the desired values in the particular places on the pdf sheet. Currently this sheet is taking a A4 size sheet formatting. However, this template doesn't include the pictures and colors. here is an example

[temp.pdf](https://github.com/rohit250992/Product_label/files/11896036/temp.pdf)


