VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExcellentFZF 
   Caption         =   "ExcellentFZF"
   ClientHeight    =   2445
   ClientLeft      =   105
   ClientTop       =   1455
   ClientWidth     =   4590
   OleObjectBlob   =   "ExcellentFZF.frx":0000
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "ExcellentFZF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private allMacros As Collection
Private isUpdating As Boolean

' Initialize the user form and populate the ComboBox with macro names
Private Sub UserForm_Initialize()
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim vbMod As VBIDE.CodeModule
    Dim line As Long
    Dim macroName As String

    ' Initialize ComboBox and master list
    Set allMacros = New Collection
    Me.ComboBox1.Clear
    Me.ListBox1.Clear

    ' Loop through all modules to collect macro names
    Set vbProj = ThisWorkbook.VBProject
    For Each vbComp In vbProj.VBComponents
        If vbComp.Type = vbext_ct_StdModule Or vbComp.Type = vbext_ct_Document Then
            Set vbMod = vbComp.CodeModule
            line = 1
            Do Until line > vbMod.CountOfLines
                macroName = vbMod.ProcOfLine(line, vbext_pk_Proc)
                ' Exclude ShowMacroSelector macro from the list
                If macroName <> "" And macroName <> "ShowExcellentFZF" Then
                    allMacros.Add macroName
                    Me.ComboBox1.AddItem macroName
                End If
                line = line + vbMod.ProcCountLines(macroName, vbext_pk_Proc)
            Loop
        End If
    Next vbComp
End Sub

' Filter the ComboBox and ListBox based on the text input in the ComboBox
Private Sub ComboBox1_Change()
    If isUpdating Then Exit Sub

    Dim i As Integer
    Dim currentText As String
    Dim filteredList As Collection
    Dim item As Variant
    Dim score As Double
    Dim minScore As Double
    Set filteredList = New Collection

    ' Get the current text in the ComboBox
    currentText = Me.ComboBox1.Text

    ' Set a minimum score for fuzzy matching
    minScore = 0.65 ' Adjust the threshold based on your requirement

    ' Filter the master list of macros using fuzzy matching
    For i = 1 To allMacros.Count
        score = JaroWinkler(currentText, allMacros(i))
        If score > minScore Then
            filteredList.Add allMacros(i)
        End If
    Next i

    ' Temporarily disable events for better performance
    isUpdating = True

    ' Update the ListBox with the filtered list
    Me.ListBox1.Clear
    For Each item In filteredList
        Me.ListBox1.AddItem item
    Next item

    ' Update the ComboBox with the filtered list
    Me.ComboBox1.Clear
    For Each item In filteredList
        Me.ComboBox1.AddItem item
    Next item

    ' Set the text back to the current text
    Me.ComboBox1.Text = currentText
    Me.ComboBox1.SelStart = Len(currentText)

    ' Ensure the UI updates
    DoEvents

    ' Simulate closing and reopening the dropdown by focusing a HiddenButton and refocusing the combobox
    Me.HiddenButton.SetFocus
    Me.ComboBox1.SetFocus
    Me.ComboBox1.DropDown

    ' Re-enable events
    isUpdating = False
End Sub


' Handle key down events for the ComboBox
Private Sub ComboBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ' Update the ComboBox text with the selected item from the ListBox
        If Me.ListBox1.ListIndex >= 0 Then
            Me.ComboBox1.Text = Me.ListBox1.List(Me.ListBox1.ListIndex)
        End If
        ' Run the selected macro
        Call RunSelectedMacro
    ElseIf KeyCode = vbKeyTab Then
        ' Cycle through the ListBox items when Tab is pressed without changing the ComboBox text
        If Me.ListBox1.ListCount > 0 Then
            If Me.ListBox1.ListIndex < 0 Then
                Me.ListBox1.ListIndex = 0
            Else
                Me.ListBox1.ListIndex = (Me.ListBox1.ListIndex + 1) Mod Me.ListBox1.ListCount
            End If
            ' Prevent the default Tab behavior
            KeyCode = 0
        End If
    ElseIf KeyCode = vbKeyEscape Then
        ' Close the form when the Escape key is pressed
        Unload Me
    End If
End Sub

' Handle key down events for the ListBox
Private Sub ListBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        ' Close the form when the Escape key is pressed
        Unload Me
    End If
End Sub

' Handle double-click events for the ListBox
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' Update the ComboBox text with the selected item and run the macro on double-click
    If Me.ListBox1.ListIndex >= 0 Then
        Me.ComboBox1.Text = Me.ListBox1.List(Me.ListBox1.ListIndex)
        Call RunSelectedMacro
    End If
End Sub

' Run the selected macro
Private Sub RunSelectedMacro()
    Dim selectedMacro As String

    ' Get the selected macro name from ListBox or ComboBox
    If Me.ListBox1.ListIndex >= 0 Then
        selectedMacro = Me.ListBox1.List(Me.ListBox1.ListIndex)
    Else
        selectedMacro = Me.ComboBox1.Text
    End If

    ' Check if a macro is selected
    If selectedMacro = "" Then
        MsgBox "Please select a macro.", vbExclamation
        Exit Sub
    End If

    ' Run the selected macro
    On Error Resume Next
    Application.run selectedMacro
    If Err.Number <> 0 Then
        MsgBox "Macro '" & selectedMacro & "' not found or failed to run.", vbExclamation
    End If
    On Error GoTo 0

    ' Close the user form
    Unload Me
End Sub

' Implementation of the JaroWinkler algo in VBA
Private Function JaroWinkler(s1 As String, s2 As String) As Double
    Dim m As Integer, t As Integer, l As Integer, s1_len As Integer, s2_len As Integer
    Dim s1_matches() As Boolean, s2_matches() As Boolean
    Dim i As Integer, j As Integer
    Dim max_dist As Integer
    Dim p As Double

    s1 = LCase(s1)
    s2 = LCase(s2)
    
    s1_len = Len(s1)
    s2_len = Len(s2)

    If s1_len = 0 Or s2_len = 0 Then
        JaroWinkler = 0
        Exit Function
    End If

    max_dist = Application.WorksheetFunction.Max(s1_len, s2_len) \ 2 - 1

    ReDim s1_matches(0 To s1_len - 1)
    ReDim s2_matches(0 To s2_len - 1)

    ' Calculate matching characters
    For i = 1 To s1_len
        For j = Application.WorksheetFunction.Max(1, i - max_dist) To Application.WorksheetFunction.Min(s2_len, i + max_dist)
            If Not s2_matches(j - 1) And Mid(s1, i, 1) = Mid(s2, j, 1) Then
                s1_matches(i - 1) = True
                s2_matches(j - 1) = True
                m = m + 1
                Exit For
            End If
        Next j
    Next i

    If m = 0 Then
        JaroWinkler = 0
        Exit Function
    End If

    ' Calculate transpositions
    j = 1
    For i = 1 To s1_len
        If s1_matches(i - 1) Then
            Do While Not s2_matches(j - 1)
                j = j + 1
            Loop
            If Mid(s1, i, 1) <> Mid(s2, j, 1) Then
                t = t + 1
            End If
            j = j + 1
        End If
    Next i
    t = t \ 2

    ' Calculate Jaro distance
    JaroWinkler = (m / s1_len + m / s2_len + (m - t) / m) / 3

    ' Calculate Jaro-Winkler distance
    l = 0
    For i = 1 To Application.WorksheetFunction.Min(4, Application.WorksheetFunction.Min(s1_len, s2_len))
        If Mid(s1, i, 1) = Mid(s2, i, 1) Then
            l = l + 1
        Else
            Exit For
        End If
    Next i

    p = 0.1 ' Scaling factor
    JaroWinkler = JaroWinkler + l * p * (1 - JaroWinkler)
End Function

