:: Place the full pathname of the file DateManager.bat 
:: in the quotes after /TR
::
:: example on my computer -  C:\Users\Nicholas\Documents\GitHub\DateManager\DateManager.bat
schtasks /Create /SC WEEKLY /D "MON, THU" /TN "Manage Dates" /TR "C:\Users\Nicholas\Documents\GitHub\DateManager\DateManager.bat" /ST 11:00
