Attribute VB_Name = "modCodeCount"
Option Explicit
Option Private Module

'Excel Add-in Code Counter
'
'Tool for counting the lines of code in Excel Add-ins (.xlam)
'Version 1.0 (July 2017)
'
'Author: Marco Krapf
'excel@marco-krapf.de
'https://marco-krapf.de/excel/
'
'License: GNU General Public License v3.0

'Variables
Dim newRibbon As IRibbonUI
Dim refresh As clsRefresh 'For holding an instance of this class
Dim coll As Collection 'Installed Add-ins
Dim selAddIn As String 'Selected Add-in
Dim eachComp As Boolean 'Counting code of each component
Dim textfile As Boolean 'Printing to text file

'Main procedure for counting lines of code
Private Sub ExcelCodeCount(selAddIn As String)
    Dim wkb As Workbook
    Dim VBCodeModule As Object
    Dim sumLines As Long, compLines As Long 'Total lines of code
    Dim sumCode As Long, compCode As Long 'Lines with active source code
    Dim sumComment As Long, compComment As Long 'Lines with comments only
    Dim sumBlank As Long, compBlank As Long 'Blank lines
    Dim i As Long, j As Long
    Dim outfile As String 'Path for the text file
    Dim channel As Integer 'Number of the output channel
    
    'Set workbook to analyse
    Set wkb = Workbooks(selAddIn)
    
    'If printing to text file is checked
    If textfile Then
        MsgBox "Please select the destination folder"
        outfile = TextFilePath() & Application.PathSeparator & wkb.Name & ".txt"
        If Left(outfile, 1) = Application.PathSeparator Then Exit Sub 'In case no folder is selected
        channel = FreeFile 'Next free output channel
        Open outfile For Output As #channel 'Open output channel
        Print #channel, "VBA PROJECT: " & UCase(wkb.Name) 'Workbook name
        Print #channel,
        Print #channel,
        Print #channel,
    End If
    
    'Loop through all components of the VBA project
    For i = 1 To wkb.VBProject.VBComponents.count
    
        compLines = 0 'Reset counter for total lines in module
        compCode = 0 'Reset counter for lines with active source code
        compComment = 0 'Reset counter for lines with comments only
        compBlank = 0 'Reset counter for empty lines
        
        Set VBCodeModule = wkb.VBProject.VBComponents(i).CodeModule

        'Loop through all lines of a VBA component
        For j = 1 To VBCodeModule.CountOfLines
            Select Case True
                Case Left(LTrim(VBCodeModule.lines(j, 1)), 3) = "Rem" _
                    Or Left(LTrim(VBCodeModule.lines(j, 1)), 1) = "'" 'Comment line
                        compComment = compComment + 1
                Case LTrim(VBCodeModule.lines(j, 1)) = "" 'Blank line
                        compBlank = compBlank + 1
                Case Else 'Active source code
                        compCode = compCode + 1
            End Select
        Next j
        
        compLines = compCode + compComment + compBlank 'Add up total lines in VBA component
        sumLines = sumLines + compLines 'Add up total lines in VBA project
        sumCode = sumCode + compCode 'Add up active scource code in VBA project
        sumComment = sumComment + compComment 'Add up lines with comments only in VBA project
        sumBlank = sumBlank + compBlank 'Add up blank lines in VBA project
        
        If eachComp Then 'When checkbox for code of each component is selected
            If textfile Then
                Print #channel, "Component: " & VBCodeModule.Name
                Print #channel, String(11 + Len(VBCodeModule), "-")
                Print #channel,
                Print #channel, "Total lines of code: " & compLines
                Print #channel,
                Print #channel, "Lines with active code: " & compCode
                Print #channel, "Lines with comments only: " & compComment
                Print #channel, "Blank lines: " & compBlank
                Print #channel,
                Print #channel,
                Print #channel,
            End If
            
            MsgBox "Component: " & VBCodeModule.Name & vbNewLine & _
                String(10, "-") & vbNewLine & vbNewLine & _
                "Total lines of code: " & compLines & vbNewLine & vbNewLine & _
                "Lines with active code: " & compCode & vbNewLine & _
                "Lines with comments only: " & compComment & vbNewLine & _
                "Blank lines: " & compBlank, , wkb.Name & " (" & VBCodeModule.Name & ")"
        End If
    Next i
    
        If textfile Then
            Print #channel, "Workbook: " & wkb.Name 'Workbook name
            Print #channel, String(10 + Len(wkb.Name), "=")
            Print #channel,
            Print #channel, "Total lines of code: " & sumLines
            Print #channel,
            Print #channel, "Lines with active code: " & sumCode
            Print #channel, "Lines with comments only: " & sumComment
            Print #channel, "Blank lines: " & sumBlank
            Close #channel 'Close output channel
        End If
          
        MsgBox "Workbook: " & wkb.Name & vbNewLine & _
            String(10, "=") & vbNewLine & vbNewLine & _
            "Total lines of code: " & sumLines & vbNewLine & vbNewLine & _
            "Lines with active code: " & sumCode & vbNewLine & _
            "Lines with comments only: " & sumComment & vbNewLine & _
            "Blank lines: " & sumBlank, , wkb.Name
            
    Set VBCodeModule = Nothing
End Sub

'Selection of the path for the text file
Private Function TextFilePath() As String
    Dim path As String
    On Error Resume Next
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Show
        path = .SelectedItems(1) 'Assign selected folder
    End With
    TextFilePath = path
    On Error GoTo 0
End Function

Public Sub RibbonRefresh()
    newRibbon.Invalidate
End Sub


'CALLBACKS
'---------

'Callback for customUI.onLoad
Public Sub CodeCounter_onLoad(ribbon As IRibbonUI)
    Set newRibbon = ribbon
    Set refresh = Nothing
    Set refresh = New clsRefresh 'Create an instance of this class
    Set refresh.App = Application
End Sub

'Callback for comboCodeCounter getItemCount
Public Sub AddIn_getItemCount(control As IRibbonControl, ByRef returnedVal)
    Dim objAddIn As AddIn
    Dim cnt As Long
    
    Set coll = Nothing
    Set coll = New Collection 'Reset collection
    selAddIn = "" 'Reset selection
    
    For Each objAddIn In Application.AddIns
        If objAddIn.Installed Then cnt = cnt + 1
    Next objAddIn
    
    returnedVal = cnt
End Sub

'Callback for comboCodeCounter getItemLabel
Public Sub AddIn_getItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
    Dim objAddIn As AddIn
    Dim cnt As Long

    For Each objAddIn In Application.AddIns
        If objAddIn.Installed Then cnt = cnt + 1
        If cnt = index + 1 Then
            returnedVal = objAddIn.Name
            coll.Add objAddIn.Name 'Add to collection
            Exit For
        End If
    Next objAddIn
    
End Sub

'Callback for comboCodeCounter onAction
Public Sub AddIn_Click(control As IRibbonControl, id As String, index As Integer)
    selAddIn = coll(index + 1) 'Assign name of the selected Add-in
End Sub

'Callback for btnCodeCounter onAction
Public Sub Count_Click(control As IRibbonControl)
    If selAddIn = "" Then
        MsgBox "Please select Add-in", vbExclamation
    Else
        If Application.Workbooks(selAddIn).VBProject.Protection = 1 Then
            MsgBox "VBAProject " & selAddIn & " is protected", vbExclamation
        Else
            Call ExcelCodeCount(selAddIn) 'Start counting code of this Add-in
        End If
    End If
End Sub

'Callback for chkComponents onAction
Public Sub Checkboxes(control As IRibbonControl, pressed As Boolean)
    Select Case control.id
        Case "chkComponents" 'Counting code of each components of a VBA project
            If pressed Then
                eachComp = True
            Else
                eachComp = False
            End If
        Case "chkTextfile" 'Printing to text file
            If pressed Then
                textfile = True
            Else
                textfile = False
            End If
    End Select
End Sub
