Attribute VB_Name = "Module1"
Public Const yourPassword As String = "your_password"
Public Const VK_SHIFT As Long = &H10

' --- WINDOWS API DECLARATIONS FOR 32-BIT AND 64-BIT ---
#If VBA7 Then
    ' For Office 64-bit and 32-bit (2010+)
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare PtrSafe Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As LongPtr, ByVal lpString2 As String) As LongPtr
#Else
    ' For Office 32-bit (2007 or older)
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Long, ByVal lpString2 As String) As Long
#End If

Private Const GHND As Long = &H42
Private Const CF_TEXT As Long = 1

Private Sub CopyToClipboard(ByVal textToCopy As String)
    #If VBA7 Then
        Dim hGlobalMemory As LongPtr, lpGlobalMemory As LongPtr
    #Else
        Dim hGlobalMemory As Long, lpGlobalMemory As Long
    #End If
    
    hGlobalMemory = GlobalAlloc(GHND, Len(textToCopy) + 1)
    If hGlobalMemory = 0 Then Exit Sub
    
    lpGlobalMemory = GlobalLock(hGlobalMemory)
    lstrcpy lpGlobalMemory, textToCopy
    GlobalUnlock hGlobalMemory
    
    If OpenClipboard(0&) <> 0 Then
        EmptyClipboard
        SetClipboardData CF_TEXT, hGlobalMemory
        CloseClipboard
    End If
End Sub

Public Sub InterceptCopy()
    If TypeName(Selection) <> "Range" Then Exit Sub

    Dim textToCopy As String
    Dim rowItem As Range, cellItem As Range
    
    If Selection.Cells.Count = 1 Then
        textToCopy = Selection.Value
    Else
        For Each rowItem In Selection.Rows
            For Each cellItem In rowItem.Cells
                textToCopy = textToCopy & cellItem.Value & vbTab
            Next cellItem
            textToCopy = Left(textToCopy, Len(textToCopy) - 1) & vbCrLf
        Next rowItem
        textToCopy = Left(textToCopy, Len(textToCopy) - 2)
    End If
    
    Call CopyToClipboard(textToCopy)
    
    Application.StatusBar = "Range copied to clipboard."
End Sub

Public Sub InterceptPaste()
    Application.EnableEvents = False
    On Error GoTo CleanUp
    
    With ActiveSheet
        .Unprotect Password:=yourPassword
        .Paste
        .Cells.Locked = True
        Selection.Locked = False
    End With

CleanUp:
    ActiveSheet.Protect Password:=yourPassword, UserInterfaceOnly:=True, _
                        AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
    Application.EnableEvents = True
End Sub

Public Sub InterceptTab()
    If GetAsyncKeyState(vbKeyShift) < 0 Then
        Application.SendKeys "{LEFT}", True
    Else
        Application.SendKeys "{RIGHT}", True
    End If
End Sub