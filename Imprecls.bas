Attribute VB_Name = "Imprecls"
Option Explicit

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" ( _
    ByVal hObject As Long, _
    ByVal nCount As Long, _
    ByRef lpObject As Any) As Long

Private Type DEVNAMES
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
End Type

Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" ( _
    ByVal lpString1 As String, _
    ByVal lpString2 As Long) As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" ( _
    ByVal lpString As Long) As Long
    
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByVal Source As Long, _
    ByVal Length As Long)
Private Declare Function GetDefaultPrinter Lib "winspool.drv" Alias "GetDefaultPrinterA" ( _
    ByVal pszBuffer As String, _
    ByRef pcchBuffer As Long) As Long

Private Declare Function SetDefaultPrinter Lib "winspool.drv" Alias "SetDefaultPrinterA" ( _
    ByVal pszPrinter As String) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const TECHNOLOGY As Long = 2

Private Function PtrToStrA(ByVal pStr As Long) As String
    Dim n As Long, s As String
    n = lstrlen(pStr)
    If n <= 0 Then Exit Function
    s = String$(n, vbNullChar)
    lstrcpy s, pStr
    PtrToStrA = Left$(s, n)
End Function


Public Function GetDefaultPrinterName() As String
    Dim n As Long, s As String

    ' Primer llamada: obtener tamaño requerido
    GetDefaultPrinter vbNullString, n
    If n <= 0 Then Exit Function

    s = String$(n, vbNullChar)

    If GetDefaultPrinter(s, n) <> 0 Then
        GetDefaultPrinterName = Left$(s, n - 1) ' sin null
    End If
End Function

Public Sub SetDefaultPrinterName(ByVal prn As String)
    If Len(prn) > 0 Then
        SetDefaultPrinter prn
    End If
End Sub
Public Function CommonDialog_SelectedPrinterDeviceName(ByVal hDevNames As Long) As String
    ' Devuelve el "DeviceName" (nombre de impresora) elegido en el ShowPrinter
    Dim p As Long
    Dim dn As DEVNAMES

    If hDevNames = 0 Then Exit Function

    p = GlobalLock(hDevNames)
    If p = 0 Then Exit Function

    ' Copiamos la estructura
    CopyMemory dn, ByVal p, LenB(dn)

    ' DeviceName es un string ANSI dentro del bloque
    CommonDialog_SelectedPrinterDeviceName = PtrToStrA(p + dn.wDeviceOffset)

    GlobalUnlock hDevNames
End Function


