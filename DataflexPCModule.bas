Attribute VB_Name = "API_DataflexPC"

' =================================================================
' DataflexPC Module for Visual Basic for Application
' =================================================================
'
' Author:       Julio L. Muller
' Version:      1.0.0
' Repository:   https://github.com/juliolmuller/VBA-Module-Dataflex
'
' =================================================================

Option Private Module
Option Explicit

'Options to set window application state
Private Enum WindowState
    MAXIMIZE = 3
    MINIMIZE = 6
    Restore = 9
End Enum

'Import library to support keyboard events
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

'Import libraries to support external application window manipulation
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'Import library to require scheduler to freeze
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Check if Dataflex is open and put it at the foreground to execute the task
Public Function IsDataflexOpen() As Boolean

    'Dimension local variables
    Dim intOpenProgram As Long
    Dim blnSucceeded   As Boolean

    'Find Dataflex window index
    intOpenProgram = FindWindow(0, "DataflexPC")
    
    'Check if Dataflex window is open
    blnSucceeded = (intOpenProgram <> 0)
    If Not (blnSucceeded) Then
        MsgBox "Please, make sure to open and log into Dataflex before runing this task."

    'Activate Dataflex window
    Else
        Call ShowWindow(intOpenProgram, WindowState.MAXIMIZE)
        Call AppActivate("DataflexPC")
        Sleep 100
    End If

    'Return success status
    IsDataflexOpen = blnSucceeded

End Function

'Run the sequence of "SendKeys" instructions
Public Function RunTaskInDataflex(arrInstructions As Variant) As Boolean

    'Dimension local variable
    Dim blnSucceeded As Boolean
    Dim i            As Integer
    Dim intVarType   As Integer

    'Loop through array and execute proper function
    blnSucceeded = True
    For i = LBound(arrInstructions) To UBound(arrInstructions)
        intVarType = varType(arrInstructions(i))
        If (intVarType = 8) Then
            If (arrInstructions(i) = Empty) Then
                SendKeys "{DEL}"
            Else
                SendKeys arrInstructions(i)
            End If
        ElseIf (intVarType = 2 Or intVarType = 3) Then
            Sleep arrInstructions(i)
        Else
            MsgBox arrInstructions(i) & " (type: " & GetVarTypeName(intVarType) & ") is not a valid instruction.", vbCritical, "Fatal Error"
            blnSucceeded = False
            Exit For
        End If
    Next i

    'Return success status
    RunTaskInDataflex = blnSucceeded

End Function

'Toggle "NUMLOCK" key. Recommended after running "SendKeys" instructions
Public Sub ToggleNumLock()

    SendKeys "{NUMLOCK}"

End Sub

'Get the variant name depending on its type
Private Function GetVarTypeName(varType As Integer) As String

    'Looks up for the anme of the variable type
    Select Case (varType)
        Case 0:
            GetVarTypeName = "empty"
        Case 1:
            GetVarTypeName = "null"
        Case 2:
            GetVarTypeName = "integer"
        Case 3:
            GetVarTypeName = "long integer"
        Case 4:
            GetVarTypeName = "single"
        Case 5:
            GetVarTypeName = "double"
        Case 6:
            GetVarTypeName = "currency"
        Case 7:
            GetVarTypeName = "date"
        Case 8:
            GetVarTypeName = "string"
        Case 9:
            GetVarTypeName = "object"
        Case 10:
            GetVarTypeName = "error value"
        Case 11:
            GetVarTypeName = "boolean"
        Case 12:
            GetVarTypeName = "variant"
        Case 13:
            GetVarTypeName = "data access object"
        Case 14:
            GetVarTypeName = "decimal value"
        Case 17:
            GetVarTypeName = "byte"
        Case 36:
            GetVarTypeName = "user defined"
        Case 8192:
            GetVarTypeName = "array"
        Case Else
            GetVarTypeName = vbNullString
    End Select

End Function
