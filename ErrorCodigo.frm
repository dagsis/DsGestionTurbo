VERSION 5.00
Begin VB.Form ErrorCodigo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   4500
      Begin VB.Timer Timer1 
         Left            =   3330
         Top             =   1185
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "PARA CONTINUAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   540
         TabIndex        =   4
         Top             =   1830
         Width           =   3390
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   1620
         TabIndex        =   3
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   135
         Left            =   1965
         TabIndex        =   2
         Top             =   840
         Width           =   90
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "CODIGO DE PRODUCTO ERRONEO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   315
         TabIndex        =   1
         Top             =   195
         Width           =   3840
      End
   End
End
Attribute VB_Name = "ErrorCodigo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
          Case 107
              Timer1.Enabled = False
              Unload Me
   End Select
End Sub

Private Sub Form_Load()
  Timer1.Interval = 1000
  Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
  Beep
End Sub
