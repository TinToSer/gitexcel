# gitExcel
Compare Excel files with Git versioning using the Microsoft SPREADSHEETCOMPARE tool.

Important: SpreadsheetCompare.exe is only available with Office Professional Plus 2013, Office Professional Plus 2016, Office Professional Plus 2019, or Microsoft 365 Apps for enterprise.

The batch scripts will help you to setup excel comparison using git diff commands.

Setup
  1. Clone or download gitExcel repository
  2. Run the "globalSetupGitExcel.bat" file (not as admin). This only needs to be run once.
  3. Run "repoSetupGitExcel.bat" for each project you want the gitExcel to work as it (creates and) modifies the .gitattributes file
  4. Use git diff: Open git console and type "git diff [Excel File Name]/hash hash", you will see SPREADSHEETCOMPARE tool opening with two excel files.
  5. For TortoiseGit:
     - Settings > Diff Viewer > Advanced > .xlsx : "X:\Path\to\gitExcel.cmd" -T %mine %base
     - Settings > Diff Viewer > Advanced > .xls : "X:\Path\to\gitExcel.cmd" -T %mine %base

----------Enjoy----------
