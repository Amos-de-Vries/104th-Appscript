# 104th-Appscript
Appscript code for 104th Menu Testing

### Menu Testing Roster Link
https://docs.google.com/spreadsheets/d/16-1DnRRN2uUHyoaJGlaT6uokMIJA_FFbh5uxS9aWvV0/edit#gid=1362002986

### Implemented Features
- New Strike System
- HTML form with a button to input information automatically based on selected cells
- Automatic input if openining form with proper selected cells
- New strike system with the option to set strikes to higher / lower values (0.5 steps) & Expirable Strikes
- Automatic qoata strike all option (strikes anyone that did not reach qoata and did not get off LOA that week, also excludes anyone who got their rank that week and didnt have the full time to get it done)
- Automatic Check Qoata strike all option (removes all qoata strikes from those who have any and met their qoata at the time of pushing the button (and those who no longere need qoata; XO+ / PVT-CPL).

### Ideas:
- Log qoata of any NCO+ of the last week, in order to look back when they cant click the strike check button in time; ie. if noone does it saterday / sunday, if done monday it would of reset the qoata and noone would have their qoata strikes removes unless the ygot their qoata in the first day.

### 0auth Permissions
- @OnlyCurrentDoc forces the authorization dialog to ask only for access to files in which the add-on or script is used, rather than all of a user's spreadsheets, documents, or forms. 
- Display and run third-party web content in prompts and sidebars inside Google applications	https://www.googleapis.com/auth/script.container.ui
- View and manage spreadsheets that this application has been installed in	https://www.googleapis.com/auth/spreadsheets.currentonly
