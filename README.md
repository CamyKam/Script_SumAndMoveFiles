# Python Script to Perform Multi-File Actions
After chatting with a friend that was complaining about how they had to open several Excel files to do calculations in each one and save the results in a new folder, I came up with this script to help her out in automating the tedious process. 

As this script was meant to be an example for her to work off of, the included file is an Ottawa Open source dataset containing information on marriage registrations each month from 2016 - Feb 2023

I created a Python script to perform the following actions in sequence:
* Look at all files starting with the word 'marriage' in the name within the folder
* For each of those files, it will sum up the amount in the 'Total' columns and group it by years found in the 'Year' column
* The results will be saved as a new file and moved to the set destination folder

**Note: This script operates with the assumption that the source files will always have the same naming structure

### Initial file
![image](https://github.com/user-attachments/assets/e1c91bed-3879-4dfb-bc8f-e44d1619795e)  ![image](https://github.com/user-attachments/assets/e625eeac-e6e3-4eb7-8c04-3429956314b4)


### File afterwards 
![image](https://github.com/user-attachments/assets/d19a181f-738f-4dc8-8eb0-352db598010e)  ![image](https://github.com/user-attachments/assets/95dbcd67-7dad-4bb2-8999-02bfaae03411)

