VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CajaVerTurnos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Turnos"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   9390
   Begin VB.CommandButton CmdMover 
      Caption         =   "|<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   135
      TabIndex        =   4
      Top             =   4965
      Width           =   840
   End
   Begin VB.CommandButton CmdMover 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   1035
      TabIndex        =   3
      Top             =   4965
      Width           =   840
   End
   Begin VB.CommandButton CmdMover 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   1965
      TabIndex        =   2
      Top             =   4965
      Width           =   840
   End
   Begin VB.CommandButton CmdMover 
      Caption         =   ">>|"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   2940
      TabIndex        =   1
      Top             =   4965
      Width           =   840
   End
   Begin MSDataGridLib.DataGrid Grid1 
      Height          =   4740
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   8361
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   2
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Saldo del Turno Abierto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5115
      TabIndex        =   6
      Top             =   4980
      Width           =   2205
   End
   Begin VB.Label LblTurno 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7380
      TabIndex        =   5
      Top             =   4980
      Width           =   1710
   End
End
Attribute VB_Name = "CajaVerTurnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cn As Connection
Private Rs2 As Recordset
Const PAGE_SIZE = 50
Dim Registro As Integer

Private WithEvents Rsd As ADODB.Recordset
Attribute Rsd.VB_VarHelpID = -1

Private Sub Form_Load()
  Conectar
  SaldoTurno
End Sub

Private Sub SaldoTurno()
  Dim cRx As ClsComprobantesL
  Dim nImporte As Double
    
  Set cRx = New ClsComprobantesL
  
  nImporte = cRx.TraerCajaAbiertaTurno()
  
  LblTurno.Caption = Format(nImporte, "#0.00")
  
End Sub

Private Sub Conectar()
  Dim sSql As String, cRsl As ClsLectura
  
  
  Dim Cadena As String
  Set cn = New ADODB.Connection
  
  cn.ConnectionString = sDb
  cn.Open
  
  Set Rsd = New Recordset
  Set Rs2 = New Recordset
  Set cRsl = New ClsLectura
  
  Dim bPermiso As Boolean
  
  bPermiso = cRsl.TraerValorDeUnCampo("UsuariosPermisos", "Valor", "Opcion=99 and Usuario=" & nUsuario)

  If bPermiso = True Then
     sSql = "SELECT * FROM Turnos ORDER BY Id Desc"
  Else
     sSql = "SELECT TOP 1 * FROM Turnos ORDER BY Id Desc"
  End If
  
  With Rsd
       .CursorLocation = adUseClient
       .PageSize = PAGE_SIZE
       .Properties("Initial Fetch Size") = PAGE_SIZE
       .Open sSql, cn, adOpenDynamic, adLockOptimistic
       .ActiveConnection = Nothing
       Set Grid1.DataSource = Paginar_Recordset(Rsd)
       CabGrid
  End With
  Registro = PAGE_SIZE
  cn.Close
End Sub

Private Sub CabGrid()
  Grid1.HeadFont.Size = 10
  Grid1.HeadFont.bold = True
  Grid1.AllowUpdate = True
  With Grid1
       .Columns(0).Visible = False
       .Columns(1).Caption = "Usuario"
       .Columns(1).Width = 1500
       .Columns(2).Caption = "Fecha"
       .Columns(2).Width = 1100
       .Columns(3).Width = 1700
       .Columns(3).Caption = "F/H Apertura"
       .Columns(4).Width = 1700
       .Columns(4).Caption = "F/H Cierre"
       .Columns(5).Width = 1000
       .Columns(5).NumberFormat = nCantDecimales
       .Columns(5).Alignment = dbgRight
       .Columns(5).Caption = "T.Turno"
  End With
End Sub

Private Function Paginar_Recordset(recset As ADODB.Recordset) As ADODB.Recordset
  
    Dim Fila As Long
    Dim Campo As Field
    Dim P As Long
  
    P = IIf(recset.AbsolutePage > 0, recset.AbsolutePage, 1)
       
    'Recordset desconectado
    With Rs2
           
        ' Si el recordset está abierto entonces lo cierra
        If .State = adStateOpen Then
            .Close
        End If
           
        ' Define y Crea los campos al recordset con Append
        For Each Campo In recset.Fields
            .Fields.Append Campo.Name, IIf(Campo.Type = 131, 6, Campo.Type), _
                           Campo.DefinedSize, Campo.Attributes
        Next Campo
           
        ' Lo Abre para agregarle los registros
           
        Rs2.Open
           
        For Fila = 1 To PAGE_SIZE
               
            ' si está en el final sale
            If recset.EOF Then
                Exit For
            End If
               
            ' va añadiendo los registros
            .AddNew
               
            ' Recorre los campos
             For Each Campo In recset.Fields
                 ' Asigna el nombre del campo
                 Rs2(Campo.Name) = Campo.Value
             Next Campo
                
             'Actualiza el recordset desconectado y el otro recordset
             .Update
             ' mueve al siguiente registro
             recset.MoveNext
  
        Next Fila
        ' Mueve al primer registro
        If Not .BOF Then
           .MoveFirst
            recset.AbsolutePage = P
        End If
        ' Establece la página actual
  
    End With
    ' devuelve a la función el conjunto de datos, es decir el recordset paginado
    Set Paginar_Recordset = Rs2
       
    ' Muestra la página actual en el caption del formulario
    Me.Caption = " Página: " & CStr(Rsd.AbsolutePage) & _
                 " de " & CStr(Rsd.PageCount)
  
End Function

