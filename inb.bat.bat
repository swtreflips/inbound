@echo off

:: Activate the Anaconda environment

call C:\Users\SilviaJulianaNavasPi\anaconda3\Scripts\activate.bat .seleniumvenv
 
:: Change to the directory where your script is located

cd C:\Users\SilviaJulianaNavasPi\OneDrive - Prime Time Packaging\Inbound Py
 
:: Run the script using the activated environment

python mainfinal4.py
 