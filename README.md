# DMLF_XTD
Script for comparisons between DMLF files of XTD tester

# -*- coding: utf-8 -*-
# Make sure that the xlsxwriter package was installed. If not:
# $cd ./Support/XlsxWriter-1.0.4/
# $sudo python setup.py install

#Example: compare 2 DMLF files:
# $python compareDMLF_XTD.py Sample_DMLF/90337BA.PR35.002.05 \
#                            Sample_DMLF/90337BA.PR35.002.06 \
#                            --hide PROMPT,RES,MIN,MAX,ARRAY_SIZE,LOG,DISPLAY,STATISTICS,ERROR,TEST \
#                            --ignore NUM,DESC \
#                            --renameFile RenameParam_90337.csv

# Compare 2 DMLF folders
# $python compareDMLF_XTD.py -d Sample_DMLF \
#                               Sample_DMLF \
#                            --dev 90337BA \
#                            --cond PR150,PR35,PR175 \
#                            --spec 002.05,002.06 \
#                            --hide PROMPT,RES,MIN,MAX,ARRAY_SIZE,LOG,DISPLAY,STATISTICS,ERROR,TEST \
#                            --ignore NUM,DESC \
#                            --renameFile RenameParam_90337.csv
