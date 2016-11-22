:: Place the name of the full path to the manage data vbs
:: script in the quotes after wscript
::
:: example on my computer - "C:\Users\Nicholas\Documents\GitHub\DateManager\manage_dates.vbs"
taskkill /IM EXCEL.EXE

wscript "C:\Users\Nicholas\Documents\GitHub\DateManager\manage_dates.vbs"
