# VBA-Searchable-Dropdown

clsSearchableDropdown

https://ExcelMacroMastery.com/

Author: Paul Kelly
Description: Searchable Dropdown list
             The code in this class allows the user to
             create a searchable dropdown for a VBA UserForm.


HOW TO USE

1. Import this Class(clsSearchableDropdown) to you project. Right-Click on the
   the workbook in the Project Window and select "Import file".

   Note: Project window is normally on the left side of the Visual Basic editor. If
   it is not visible select View->Project Explorer(Ctrl + R) from the menu.

2. Insert a UserForm if you dont already have one and a listbox and textbox to your UserForm

Place the textbox where you want it to appear and add the font size and settings that you require.

The listbox can be placed anywhere as the code will resize and set the position and font based on the 
textbox.

3. Add the code below in the UserForm:

``` 
Private oEventHandler As New clsSearchableDropdown

Public Property Let ListData(ByVal rg As Range)
  oEventHandler.ItemsRange = rg.Value
End Property

Private Sub UserForm_Initialize()

  ' Settings
  With oEventHandler

    ' Attach the textbox and listbox to the class
     Set .SearchListBox = Me.ListBox1
     Set .SearchTextBox = Me.TextBox1

    ' Default settings
    .MaxRows = 6
    .ShowAllMatches = False
    .CompareMethod = vbTextCompare
    .WindowsVersion = True

  End With

End Sub

Private Sub UserForm_Terminate()
    Set oEventHandler = Nothing
End Sub
```

4. Add the following code to the module that will display the UserForm

   Note: The range should be the range of the data that you want to filter
   on the form
``` 
    Sub Main()

      Dim frm As UserForm1
      Set frm = UserForms.Add(UserForm1.Name)
      frm.ListData = Sheet1.Range("A1").CurrentRegion

      frm.show
    
    End Sub
``` 


















