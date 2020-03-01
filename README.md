## Activity Tracker
#### a simple activity tracker built with Excel and VBA to track time spent on investigating issues raised via e-mail

- The tracker loops through 3 mailboxes and pulls e-mail messages from the last 4 days (current date to -3 days) from *Inbox* only. 
- Information such as *Sender Name*, *Body* (first 200 characters with empty rows converted to tabs for better readability), *Received Time*, etc. is saved into the *adhoc_tracker.xlsm* file.
- Only unique e-mails by *Sender Name* and *Subject* are added to the list.
- Tasks can be assigned to employees by inserting 'x' into specific columns (*Emp1-5*).
- Only first three assigned activities can be monitored (copy macros, adjust where needed, and add more buttons if you wish - below is a hint)
```
Sub Track_4()
(...)
If IsEmpty(ActiveSheet.Range("A5").Value) = False Then
    If ActiveSheet.Cells(r - 1, 12).Value <> ActiveSheet.Range("E5").Value...
(...)
        ActiveSheet.Cells(r, 8).Value = ActiveSheet.Range("A5").Value
        ActiveSheet.Cells(r, 9).Value = ActiveSheet.Range("B5")...
        (...)
```
***Psst... you could use Caller property to have only one Start/Stop macro that would work differently based on a button's position***
- Every employee can refresh their tracker (*at_employee1-3.xlsm*) by hitting the **[Refresh Tasks]** button. When they finish working on the assigned tasks, they can export log to the main file (*adhoc_tracker.xlsm*) by hitting the **[Export Data]** button.
- You can't start an activity if the *Category L1* column remains empty.
- You can start working on another activity without having to stop the current one. The activity you started previously will end and a new row will be created for another activity at the same time.
- Only activities that have the *End Time* column not empty will be exported to the main file.

Note that all files can be saved in and used from any location. Please only remember to change path to the master file in the code accordingly for all individual employees' files: 
```
    Set objWbk = Workbooks.Open("C:\Users\DOM\Downloads\adhoc_tracker.xlsm")
```

If you wish, you can change the number of mailboxes to loop through by changing the array within the VBA script as well. Find this line in the code: 
```
mArr = Array("SH-Derivatives@mycompany.com", "SH-Banking@mycompany.com", "robert.lendzion@mycompany.com")
```
There are also regular expressions you may want to change as well:
```
With rgx
    .Pattern = ".\d{8}$"
```
and
```
With rgxm
    .Pattern = "[a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+"
```

These are just templates and can by modified in ANY way, to meet your needs.
