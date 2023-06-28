# Product_label
# Read and store content of an excel file.
 This part of the script reads your excel file and store all its contents. The path must be provided by the user to read the specific file.

# Write the dataframe object into csv file.
 Now to read the values of the file correctly and not just the formulas present in the cell. I have converted it into a CSV and again saved    it as an excel, so I can get the desired result. 

# Load the entire workbook & # Load one worksheet.
 This section of the script open ups the whole excel file as a workbook and BPR as a worksheet. Now, I saved all the values required to make the shipping label like Name, Daily Dose, capsule per bottle, lot number, manufacturing date and customer id for the worksheet. 

## Making supplement chart 
Now to make an example like supplement chart I have just used the Item name, percent and dosage from the formula sheet and put all the values in a table form. 
╒════╤═══════════════════════════════════════════════════════════════════╤═══════════╤══════════╕
│    │ Ingredient                                                        │   percent │   dosage │
╞════╪═══════════════════════════════════════════════════════════════════╪═══════════╪══════════╡
│  0 │ 5-Hydroxytryptophan                                               │     1     │    100   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│  1 │ SensorilÂ® Ashwagandha Extract (10% Withanolides) (root and leaf) │     1     │    135   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│  2 │ Bacopa                                                            │     1     │    300   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│  3 │ Bifidobacterium bifidum                                           │     1     │      4   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│  4 │ Citrus Bioflavonoids                                              │     1     │    100   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│  5 │ Cordyceps sinensis (mycelium)                                     │     1     │    500   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│  6 │ L-5-methyltetrahydrofolate                                        │     0.95  │    200   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│  7 │ Ginkgo Biloba Extract                                             │     1     │    120   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│  8 │ Panax Ginseng                                                     │     1     │    300   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│  9 │ Green Tea Phytosome (19% Polyphenols, 13% EGCG) (leaf)            │     1     │    120   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│ 10 │ L-Glutamine                                                       │     1     │    500   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│ 11 │ L-Glutathione (Reduced)                                           │     1     │    250   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│ 12 │ Lactobacillus rhamnosus                                           │     1     │      5   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│ 13 │ Magnesium (Hydroxide)                                             │     0.33  │     50   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│ 14 │ Milk Thistle                                                      │     1     │    250   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│ 15 │ N-Acetyl-L-Cysteine                                               │     1     │    250   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│ 16 │ Phosphatidylserine 70% (sunflower)                                │     1     │    100   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│ 17 │ Resveratrol                                                       │     1     │    250   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│ 18 │ Rhodiola Rosea Extract                                            │     1     │    300   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│ 19 │ Vitamin B1 (Thiamine HCl)                                         │     0.98  │      5   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│ 20 │ Vitamin B12 (Hydroxycobalamin)                                    │     0.01  │     40   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│ 21 │ Vitamin B2 (Riboflavin)                                           │     0.98  │     10   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│ 22 │ Vitamin B5 (Calcium-D-Pantothenate)                               │     0.9   │     10   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│ 23 │ Vitamin B6 (Pyridoxal-5-Phosphate)                                │     0.63  │      8.5 │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│ 24 │ Vitamin C (Ascorbic acid)                                         │     0.99  │    350   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│ 25 │ Vitamin D3 (Cholecalciferol) (vegan)                              │     1     │     25   │
├────┼───────────────────────────────────────────────────────────────────┼───────────┼──────────┤
│ 26 │ Zinc (Citrate)                                                    │     0.313 │     10   │

#Creating a Template.

This section of the code is creating a template of the suggested shipping label. It inserts all the desired values in the particular places on the pdf sheet. Currently this sheet is taking a A4 size sheet formatting. However, this template doesn't include the pictures and colors. here is an example

[temp.pdf](https://github.com/rohit250992/Product_label/files/11896036/temp.pdf)


