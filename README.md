# DMLF_XTD
Script for comparisons between DMLF files of XTD tester

The compiled script can be found in the distributions (**dist**) folder.

Templates of compare table and rename table can be found in the **Support** folder

For detail information:

    $./dist/pyCompareXTDDmlf -h

# Command line example

Example 1: compare 2 DMLF files:

    $./dist/pyCompareXTDDmlf   Sample_DMLF/90337BA.PR35.002.05 \
                               Sample_DMLF/90337BA.PR35.002.06 \
                               --hide PROMPT,RES,MIN,MAX,ARRAY_SIZE,LOG,DISPLAY,STATISTICS,ERROR,TEST \
                               --ignore ERROR,PROMPT,TEST \
                               --renameFile RenameParam_90337.csv

Example 2: compare DMLF files in 2 folders (1 device, 2 spec versions):

    Folder1_path/90337BA.PR150.002.05   Vs  Folder2_path/90337BA.PR150.002.06
    Folder1_path/90337BA.PR35.002.05    Vs  Folder2_path/90337BA.PR35.002.06
    Folder1_path/90337BA.PR175.002.05   Vs  Folder2_path/90337BA.PR175.002.06

    $./dist/pyCompareXTDDmlf -D Folder1_path \
                                Folder2_path \
                               --dev 90337BA \
                               --cond PR150,PR35,PR175 \
                               --spec 002.05,002.06 \
                               --hide PROMPT,RES,MIN,MAX,ARRAY_SIZE,LOG,DISPLAY,STATISTICS,ERROR,TEST \
                               --ignore ERROR,PROMPT,TEST \
                               --renameFile RenameParam_90337.csv

Example 3: compare DMLF files in 2 folders (2 devices, 2 spec versions):

    Folder1_path/90337BA.PR150.002.05   Vs  Folder2_path/90337CA.PR150.002.06
    Folder1_path/90337BA.PR35.002.05    Vs  Folder2_path/90337CA.PR35.002.06
    Folder1_path/90337BA.PR175.002.05   Vs  Folder2_path/90337CA.PR175.002.06

    $./dist/pyCompareXTDDmlf   -D Folder1_path
                                  Folder2_path
                               --dev  90337BA,90337CA
                               --cond PR150,PR35,PR175
                               --spec 002.05,002.06
                               --hide PROMPT,RES,MIN,MAX,ARRAY_SIZE,LOG,DISPLAY,STATISTICS,ERROR,TEST
                               --ignore ERROR,PROMPT,TEST
                               --renameFile RenameParam_90337.csv