VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_Open()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect Password:=yourPassword
        ws.Cells.Locked = True
        ws.Protect Password:=yourPassword, UserInterfaceOnly:=True, _
                   AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
    Next ws
    
    On Error Resume Next
    ActiveCell.Locked = False
    On Error GoTo 0
    
    Application.OnKey "^c", "InterceptCopy"
    Application.OnKey "^v", "InterceptPaste"
    Application.OnKey "{TAB}", "InterceptTab"
    Application.OnKey "+{TAB}", "InterceptTab"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.OnKey "^c"
    Application.OnKey "^v"
    Application.OnKey "{TAB}"
    Application.OnKey "+{TAB}"
    Application.StatusBar = False
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Application.StatusBar <> False Then
        Application.StatusBar = False
    End If
    
    Application.EnableEvents = False
    Application.Cursor = xlDefault
    
    Sh.Unprotect Password:=yourPassword
    Sh.Cells.Locked = True
    Target.Locked = False
    Sh.Protect Password:=yourPassword, UserInterfaceOnly:=True, _
                AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
    
    Application.Cursor = xlDefault
    Application.EnableEvents = True
End Sub