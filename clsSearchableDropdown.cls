VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSearchableDropdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' clsSearchableDropdown
' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' Version: 2.1 Updated 30-01-2023
' Updates in this version:
' * Replaced the two filtering methods that used the VBA filter function(used for Mac)
'   and ArrayList(for windows). Arraylist is now problematic as it requires installation
'   of a specific .Net Library.
'   The updated FilterListBox method uses the array to filter and is 10 times
'   faster than the VBA function. It will work on both Windows and Mac.
'
' Description: Searchable Dropdown list
'              The code in this class allows the user to
'              create a searchable dropdown list of items.
'
'
' HOW TO USE
'
' 1. Import this Class(clsSearchableDropdown) to you project. Right-Click on the
'    the workbook in the Project Window and select "Import file".
'
'    Note: Project window is normally on the left side of the Visual Basic editor. If
'    it is not visible select View->Project Explorer(Ctrl + R) from the menu.
'
' 2. Insert a UserForm if you don't already have one
'
' 3. Add a listbox and textbox to your UserForm
'
' 4. Add the code below in the UserForm
'
'    *** START OF CODE TO ADD ***
'    Private oEventHandler As New clsSearchableDropdown
'
'    Public Property Let ListData(ByVal rg As Range)
'        oEventHandler.List = rg.Value
'    End Property
'
'    Private Sub UserForm_Initialize()
'
'       With oEventHandler
'           Set .SearchListBox = Me.ListBox1
'           Set .SearchTextBox = Me.TextBox1
'
'        ' Settings
'           .MaxRows = 6 ' set the number of items to be displayed
'           .ShowAllMatches = False ' to show all the matches
'           .CompareMethod = vbTextCompare ' use vbBinaryCompare for case sensitivity
'       End With
'
'    End Sub
'
'    Private Sub UserForm_Terminate()
'       Set oEventHandler = Nothing
'    End Sub
'
'    *** END OF CODE TO ADD ***
'
' 5. Add the following code to the module that will display the UserForm. If you already have
'    code to display the UserForm then you only need the ListData line.
'
'    Note: The range should be the range of the data that you want to filter
'    on the form
'
'     *** START OF CODE TO ADD ***
'    Dim frm As UserForm1
'    Set frm = UserForms.Add(UserForm1.name)
'    frm.ListData = Sheet1.Range("A1").CurrentRegion
'
'    frm.Show
'
'     *** END OF CODE TO ADD ***
'
'
'  LICENSE ETC.
'  You can use this code free in any of your software applications and distribute as you wish.
'  Please do not alter or remove the authors name or website.


Private Const m_conMaxRows As Long = 6

' Configurable Settings
Private m_compareMethod As VbCompareMethod  ' Determines case sensitivity in the search
Private m_listOfItems As Variant            ' This is the array of items that is filtered
Private m_maxRows As Long                   ' The number of rows to be displayed in the listbox.
Private m_startText As String               ' Start text in the textbox
Private m_showAllMatches As Boolean         ' True: shows all matches. False show the number of rows specified by m_MaxRows

' This is used to prevent events running when changes are being made
Private m_UpdateControl As Boolean


Private m_textboxStartingState As Boolean   ' Used to decide when to remove the starting text

Private WithEvents myListBox As MSForms.listbox
Attribute myListBox.VB_VarHelpID = -1
Private WithEvents myTextBox As MSForms.TextBox
Attribute myTextBox.VB_VarHelpID = -1

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU


' PROPERTIES
Public Property Let CompareMethod(ByVal value As VbCompareMethod)
    m_compareMethod = value
    Call FilterListBox
End Property
Public Property Get SelectedItem() As String
    SelectedItem = IIf(m_textboxStartingState = True, "", myTextBox.value)
End Property
Public Property Let List(ByVal value As Variant)
    m_listOfItems = value
End Property
' The number of rows that will be visible in the listbox
Public Property Let MaxRows(ByVal value As Long)
    m_maxRows = value
    Call FilterListBox
End Property
' Set the text to be displayed in the textbox before the search
Public Property Let StartText(ByVal text As String)
    m_startText = text
    SetTextboxValue (m_startText)
End Property
' If true include all matches in the listbox. If false only show the
' rows specified by m_MaxRows
Public Property Let ShowAllMatches(ByVal state As Boolean)
    m_showAllMatches = state
    Call FilterListBox
End Property

Public Property Set SearchListBox(ByVal oListBox As MSForms.listbox)
    Set myListBox = oListBox
    Call InitializeListBox
End Property

Public Property Set SearchTextBox(ByVal oTextBox As MSForms.TextBox)
    Set myTextBox = oTextBox
    Call InitializeTextBox
End Property

' CLASS EVENTS

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
' Description: Set the defaults
Private Sub Class_Initialize()

    m_compareMethod = vbTextCompare
    m_maxRows = m_conMaxRows
    m_startText = "Type the item you wish to search for"
    m_showAllMatches = False
    
End Sub


' LISTBOX EVENTS

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
Private Sub myListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' If the user clicks or presses enter then
    ' place the selected value in the textbox
    If m_UpdateControl = False Then
        SetTextboxValue myListBox.value
        Call ShowListbox(False)
    End If
End Sub

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
Private Sub myListBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' When the key is down in the listbox turn on
    ' m_UpdateControl to prevent the click event occurring
    If KeyCode = vbKeyDown Then
        m_UpdateControl = True
    ElseIf KeyCode = vbKeyUp Then
        m_UpdateControl = True
        CheckListBoxFirstItem
    ElseIf KeyCode = vbKeyReturn Then
        ' swallow the enter keycode as it passes on to the ok button
        KeyCode = 0
        SetTextboxValue myListBox.value
        Call ShowListbox(False)
    End If
End Sub

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
Private Sub myListBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Turn update control off - turned on in KeyDown
    If KeyCode = vbKeyDown Then
        m_UpdateControl = False
    ElseIf KeyCode = vbKeyUp Then
        m_UpdateControl = False
    End If
End Sub

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
Private Sub List_MoveDown()
    m_UpdateControl = True
    If m_textboxStartingState = False Then
        With myListBox
            .SetFocus
            If .ListIndex < .ListCount - 1 Then
                .ListIndex = .ListIndex + 1
                .Selected(.ListIndex) = True
            End If
        End With
    End If
    m_UpdateControl = False
End Sub


' LISTBOX SUBS
Private Sub UpdateListbox(items As Variant)
 
    With myListBox
  
        ' Reload listbox
        .Clear
        .ForeColor = rgbBlack
        
        ' Set the listbox size
        If IsEmpty(items) Then
            ' No items found
            .List = Array("No items found!")
            .ForeColor = rgbRed
        Else
            ' items found
            .List = items
            .ListIndex = 0
        End If
        
        ' Resize the listbox
        Call SetListboxPosition
         
         ' If show all matches then have a scrollbar
        If m_showAllMatches = True Then
            Call MakeAllMatchesAvailable
        Else
            .Height = ResizeListbox(myListBox, myTextBox.Font.Size)
        End If
    
    End With

End Sub


Private Sub MakeAllMatchesAvailable()

    With myListBox
     
        ' To get the scrollbar working correctly it is necessary to
        ' turn IntegralHeight off and on
        .IntegralHeight = False
        .Height = ResizeListbox(myListBox, myTextBox.Font.Size)
        .IntegralHeight = True

        ' List index will not highlight to first unless the second
        ' one is highlighted first. It might be to do with the resizing
        ' from the Integral height
        If .ListCount > 1 Then .ListIndex = 1
        .ListIndex = 0

    End With

End Sub

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
Private Sub InitializeListBox()
    ' Remove any automatic resizing of the listbox
    myListBox.IntegralHeight = False
End Sub

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
Public Sub CheckListBoxFirstItem()
    If myListBox.ListIndex = 0 Then
        m_UpdateControl = False
        SelectTextBox
    End If
End Sub

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
Private Function ShowListbox(Optional ByVal show As Boolean = True)
    myListBox.Visible = show
End Function

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
Private Function SetListboxPosition()
    
    With myListBox
        .Left = myTextBox.Left
        .Top = myTextBox.Top + (myTextBox.Height)
        .Width = myTextBox.Width
         Call ShowListbox(True)
    End With
    
End Function

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
Private Function ResizeListbox(myListBox As MSForms.listbox, fontSize As Double) As Double

    ' Set listbox font to the same size as the textbox
    myListBox.Font.Size = fontSize

    Dim ItemCount As Long
    ItemCount = IIf(myListBox.ListCount > m_maxRows, m_maxRows, myListBox.ListCount)

    Dim itemSize As Double
    
    ' the font size is itself plus a quarter for the space between rows
    itemSize = myListBox.Font.Size + (myListBox.Font.Size / 4)
    
    ' Font 10 has different sizing
    Dim extraspace As Double
    If fontSize = 10 Then
        extraspace = 4
    Else
        ' If 2 or less items then the listbox news to be taller
        If myListBox.ListCount <= 2 Then
            extraspace = 3
        Else
            extraspace = 2
        End If
    End If

    ResizeListbox = (itemSize * ItemCount) + extraspace
    
End Function


' TEXTBOX EVENTS

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
Private Sub myTextBox_Change()

    If m_UpdateControl = False Then
        If Trim(myTextBox.value) = "" Then
            Call InitializeTextBox
        Else
            If m_textboxStartingState = True Then
                m_textboxStartingState = False
                Call RemoveStartingText
            End If
            Call FilterListBox
        End If
    End If
    
End Sub

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
Private Sub myTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    ' If the user presses keyup/down or enter
    ' while in the textbox
    If KeyCode = vbKeyDown Then
        List_MoveDown
    ElseIf KeyCode = vbKeyReturn Then
        ' swallow the enter keycode as it passes on to the ok button
        KeyCode = 0
        If IsNull(myListBox.value) = False And m_textboxStartingState = False Then
            SetTextboxValue myListBox.value
        End If
        Call ShowListbox(False)
    ElseIf KeyCode = vbKeyEscape Then
        ' swallow the esc keycode
        Call InitializeTextBox
        KeyCode = 0
    End If

End Sub


' TEXTBOX SUBS

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
Private Sub SetTextboxValue(ByVal text As String)

    If m_UpdateControl = False Then
        
        With myListBox
            m_UpdateControl = True
            ' Set the listbox selected value to the textbox
            ' and hide the listbox
            myTextBox.value = text
            
            Call SelectTextBox
            
            m_UpdateControl = False

        End With
        
    End If
    
End Sub

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
Private Sub InitializeTextBox()
    ' Set the starting text and position
    m_textboxStartingState = True
    SetTextboxValue m_startText
    myTextBox.SelStart = 0
    myTextBox.ForeColor = rgbBlue
    
    Call ShowListbox(False)
End Sub

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
' Description:  Replace the default starting text with the letter the user has typed.
'               This will work if the user types at any position in the starting text.
Private Sub RemoveStartingText()
    
    m_UpdateControl = True

    With myTextBox
                
        .text = Mid(.value, .SelStart, 1)
        .ForeColor = rgbBlack
    End With
    
    m_UpdateControl = False
    
End Sub


' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
Private Function SelectTextBox()
    With myTextBox
        '.SelStart = 0
        .SetFocus
    End With
End Function


' GENERAL SUBS

' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' YouTube video: https://youtu.be/gkLB-xu_JTU
' Filters the Listbox
Public Sub FilterListBox()

    ' In case a setting like MaxRows is used before the
    ' range is set
    If IsEmpty(m_listOfItems) Then Exit Sub

    m_UpdateControl = True

    Dim items As Variant
    items = FilterData
      
    Call UpdateListbox(items)
      
    m_UpdateControl = False
    
End Sub

Private Function FilterData() As Variant
   
    Dim textPattern As String
    textPattern = myTextBox.value
    
    ' Create an array to store the filtered items
    Dim filteredItems() As String
    ReDim filteredItems(0 To m_maxRows - 1)

    ' Read through all the items in the full list
    Dim i As Long
    Dim count As Long: count = 0
    For i = LBound(m_listOfItems) To UBound(m_listOfItems)
        ' Using Instr instead of Like so we can set the case sensitivity
        If InStr(1, m_listOfItems(i, 1), textPattern, m_compareMethod) > 0 Then
            filteredItems(count) = m_listOfItems(i, 1)
            count = count + 1
            If m_showAllMatches = False Then
                ' Only show the max number of rows
                If count >= m_maxRows Then Exit For
            End If
        End If
    Next
    
    ' use variant so we can check later if the array is empty
    Dim finalItems As Variant
    If count > 0 Then
        ReDim finalItems(0 To count - 1)
        For i = 0 To count - 1
            finalItems(i) = filteredItems(i)
        Next i
    End If
    
    FilterData = finalItems

End Function

