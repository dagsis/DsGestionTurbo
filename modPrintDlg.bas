Attribute VB_Name = "modPrintDlg"
Option Explicit

' ===== WinAPI PrintDlg =====
Private Type PrintDlg
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hDC As Long
    Flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As Long
    lpSetupTemplateName As Long
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Private Type DEVNAMES
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
End Type
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long

Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (ByRef pPD As PrintDlg) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As Long, ByVal Length As Long)

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

' Flags PrintDlg
Public Const PD_ALLPAGES As Long = &H0&
Public Const PD_SELECTION As Long = &H1&
Public Const PD_PAGENUMS As Long = &H2&
Public Const PD_NOSELECTION As Long = &H4&
Public Const PD_NOPAGENUMS As Long = &H8&
Public Const PD_COLLATE As Long = &H10&
Public Const PD_PRINTTOFILE As Long = &H20&
Public Const PD_PRINTSETUP As Long = &H40&
Public Const PD_NOWARNING As Long = &H80&
Public Const PD_RETURNDC As Long = &H100&
Public Const PD_RETURNIC As Long = &H200&
Public Const PD_RETURNDEFAULT As Long = &H400&
Public Const PD_SHOWHELP As Long = &H800&
Public Const PD_ENABLEPRINTHOOK As Long = &H1000&
Public Const PD_ENABLESETUPHOOK As Long = &H2000&
Public Const PD_ENABLEPRINTTEMPLATE As Long = &H4000&
Public Const PD_ENABLESETUPTEMPLATE As Long = &H8000&
Public Const PD_ENABLEPRINTTEMPLATEHANDLE As Long = &H10000
Public Const PD_ENABLESETUPTEMPLATEHANDLE As Long = &H20000
Public Const PD_USEDEVMODECOPIES As Long = &H40000
Public Const PD_USEDEVMODECOPIESANDCOLLATE As Long = &H40000
Public Const PD_DISABLEPRINTTOFILE As Long = &H80000
Public Const PD_HIDEPRINTTOFILE As Long = &H100000
Public Const PD_NONETWORKBUTTON As Long = &H200000

Private Function PtrToStrA(ByVal pStr As Long) As String
    Dim n As Long, s As String
    n = lstrlen(pStr)
    If n <= 0 Then Exit Function
    s = String$(n, vbNullChar)
    lstrcpy s, pStr
    PtrToStrA = Left$(s, n)
End Function

Private Function GetDeviceNameFromDevNames(ByVal hDevNames As Long) As String
    Dim p As Long
    Dim dn As DEVNAMES

    If hDevNames = 0 Then Exit Function

    p = GlobalLock(hDevNames)
    If p = 0 Then Exit Function

    CopyMemory dn, p, LenB(dn)
    GetDeviceNameFromDevNames = PtrToStrA(p + dn.wDeviceOffset)

    GlobalUnlock hDevNames
End Function

' ====== API pública: muestra diálogo y devuelve datos ======
Public Function ShowWindowsPrintDialog( _
    ByVal ownerHwnd As Long, _
    ByRef outPrinterName As String, _
    ByRef outFromPage As Long, _
    ByRef outToPage As Long, _
    ByRef outCopies As Long, _
    ByRef outCollate As Boolean, _
    ByRef outUsePageRange As Boolean) As Boolean

    Dim pd As PrintDlg

    pd.lStructSize = Len(pd)              ' en VB6 Len y LenB funcionan, Len está ok
    pd.hwndOwner = ownerHwnd
    pd.hInstance = App.hInstance          ' CLAVE: evita que falle “silencioso” en muchas PCs

    pd.nMinPage = 1
    pd.nMaxPage = 32000
    pd.nFromPage = 1
    pd.nToPage = 1

    ' Flags recomendados
    pd.Flags = PD_ALLPAGES Or PD_NOSELECTION Or PD_RETURNDC Or PD_USEDEVMODECOPIESANDCOLLATE


    If PrintDlg(pd) = 0 Then
        Dim e As Long
        e = CommDlgExtendedError()

        If e <> 0 Then
            MsgBox "PrintDlg NO mostró el diálogo. CommDlgExtendedError=" & e, vbCritical, "Impresión"
        End If

        ShowWindowsPrintDialog = False
        Exit Function
    End If

    outPrinterName = GetDeviceNameFromDevNames(pd.hDevNames)
    outFromPage = pd.nFromPage
    outToPage = pd.nToPage
    outCopies = pd.nCopies
    outCollate = ((pd.Flags And PD_COLLATE) = PD_COLLATE)
    outUsePageRange = ((pd.Flags And PD_PAGENUMS) = PD_PAGENUMS)

    ' liberar handles asignados por PrintDlg (importante)
    If pd.hDevMode <> 0 Then GlobalFree pd.hDevMode
    If pd.hDevNames <> 0 Then GlobalFree pd.hDevNames

    ShowWindowsPrintDialog = (Len(outPrinterName) > 0)
End Function


