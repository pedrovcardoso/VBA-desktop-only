Attribute VB_Name = "Módulo1"
Public Const minhaSenha As String = "sua_senha"
Public Const VK_SHIFT As Long = &H10

' --- DECLARAÇÕES DA API DO WINDOWS PARA 32-BIT E 64-BIT ---
#If VBA7 Then
    ' Para Office 64-bit e 32-bit (versão 2010 e mais recentes)
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
    ' Para Office 32-bit (versão 2007 e mais antigas)
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

Private Sub CopiarParaClipboard(ByVal textoParaCopiar As String)
    #If VBA7 Then
        Dim hGlobalMemory As LongPtr, lpGlobalMemory As LongPtr
    #Else
        Dim hGlobalMemory As Long, lpGlobalMemory As Long
    #End If
    
    hGlobalMemory = GlobalAlloc(GHND, Len(textoParaCopiar) + 1)
    If hGlobalMemory = 0 Then Exit Sub
    
    lpGlobalMemory = GlobalLock(hGlobalMemory)
    lstrcpy lpGlobalMemory, textoParaCopiar
    GlobalUnlock hGlobalMemory
    
    If OpenClipboard(0&) <> 0 Then
        EmptyClipboard
        SetClipboardData CF_TEXT, hGlobalMemory
        CloseClipboard
    End If
End Sub

Public Sub InterceptarCopia()
    If TypeName(Selection) <> "Range" Then Exit Sub

    Dim textoParaCopiar As String
    Dim linha As Range, celula As Range
    
    If Selection.Cells.Count = 1 Then
        textoParaCopiar = Selection.Value
    Else
        For Each linha In Selection.Rows
            For Each celula In linha.Cells
                textoParaCopiar = textoParaCopiar & celula.Value & vbTab
            Next celula
            textoParaCopiar = Left(textoParaCopiar, Len(textoParaCopiar) - 1) & vbCrLf
        Next linha
        textoParaCopiar = Left(textoParaCopiar, Len(textoParaCopiar) - 2)
    End If
    
    Call CopiarParaClipboard(textoParaCopiar)
    
    Application.StatusBar = "Intervalo copiado para a área de transferência."
End Sub

Public Sub InterceptarColar()
    Application.EnableEvents = False
    On Error GoTo Limpar
    
    With ActiveSheet
        .Unprotect Password:=minhaSenha
        .Paste
        .Cells.Locked = True
        Selection.Locked = False
    End With

Limpar:
    ActiveSheet.Protect Password:=minhaSenha, UserInterfaceOnly:=True, _
                        AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
    Application.EnableEvents = True
End Sub

Public Sub InterceptarTab()
    If GetAsyncKeyState(vbKeyShift) < 0 Then
        Application.SendKeys "{LEFT}", True
    Else
        Application.SendKeys "{RIGHT}", True
    End If
End Sub
