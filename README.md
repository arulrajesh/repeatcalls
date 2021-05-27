## Repeat Calls
Macro workbook for Repeat Call calculator for Excel.
## Instructions
This repositiory contains the file [Repeat_calls.xlsm](https://github.com/arulrajesh/repeatcalls/raw/dev/Repeat_calls.xlsm).
Download the workbook and follow the instructions in the workbook.
- Based on your security settings in excel, you might be prompted for "enable editing" and "enable content". You can safely enable them.

1. Double click the cell for a file open dialog box. You can click this box even if the file is already open, this will populate the dropdown cells below to reflect all the sheets in the workbook. If the file is not Already Open the file will be opened and the dropdown list will be populated in cell B2.
![step 1](https://github.com/arulrajesh/repeatcalls/blob/dev/images/Capture1.PNG)
1. Select the file and click open. *I am selecting "Book1" for the purpose of this demo, your filename might be different.* The cell "B2" will be updated with a dropdown menu which can be used to select the sheet which has relevant data.
![step 2](https://github.com/arulrajesh/repeatcalls/blob/dev/images/capture4.png)
  *This is what the sample data looks like in Book1.xlsx*
  ![step 2a](https://github.com/arulrajesh/repeatcalls/blob/dev/images/capture7.png)
1. Once the SheetName is selected, the dropdown menus in cell B3 and B4 will updated for the UniqueID and DateTime.
![step 3](https://github.com/arulrajesh/repeatcalls/blob/dev/images/capture3.png)
1. Select Unique ID.
1. Select DateTime.
1. The Repeat Buckets can be entered as numbers seperated by commas eg: 1,3,5,7.
   1. Paste Values for formulas does exactly what the name suggests. Default is to leave it as TRUE for faster processing. Click on "Double click to calculate repeat calls"
![step 4](https://github.com/arulrajesh/repeatcalls/blob/dev/images/capture6.png)
1. If there were no errors you should now be taken to the workbook with all the repeat calls for the buckets in seperate columns with a repeat call denoted by a 1.
![step 5](https://github.com/arulrajesh/repeatcalls/blob/dev/images/capture8.png)

## Vishesh Tippini
- All the code and error handling can be seen by presseing Alt+F11 in the excel workbook. Its all in module1 of the VBA project.
- There is some error handling. Is there a better way to do this?

## Contributing
- Suggestions welcome