# matrix_to_table
Simple scripts that are rewriting matrices into tables.

## indexing_table.R
Is a script design to prepare indexing tables (eg. for illumina sequencing), from indexing matrices. Especially useful when preparing amplicon libraries using custom design indexes.

**Input file:**

Excel file (**.xls** or **.xlsx**) with one or multiple sheets. Number and position of indexing matrices does not matter. However the matrix layout is important, especially location of name **"plate"**. Check example file, **Greenland_example.xlsx**.

Indexes matrices should looks like this:

plate | i7 | index7_1 | index7_2 | index7_3
------|----|----------|----------|---------
**i5** | plate_name | 1 | 2 | 3
**index5_1** | A | sample0289 | sample0297 | sample0305
**index5_2** | B | sample0290 | sample0298 | sample0306
**index5_3** | C | sample0291 | sample0299 | sample0307

Output file will get prefix **"MiSeq_"** and should look like this:

i5 | i7 | samples | plate
---|----|---------|------
index5_1 | index7_1 | sample0289 | plate_name
index5_2 | index7_1 | sample0290 | plate_name
index5_3 | index7_1 | sample0291 | plate_name
index5_1 | index7_2 | sample0297 | plate_name
index5_2 | index7_2 | sample0298 | plate_name
index5_3 | index7_2 | sample0299 | plate_name
index5_1 | index7_3 | sample0305 | plate_name
index5_2 | index7_3 | sample0306 | plate_name
index5_3 | index7_3 | sample0307 | plate_name
