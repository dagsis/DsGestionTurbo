VERSION 5.00
Begin VB.Form FrmAdicion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADICION"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   1545
   ClientWidth     =   5100
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   5100
   Begin VB.TextBox TxtMesa 
      Height          =   285
      Left            =   735
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   540
   End
   Begin VB.Frame Frame2 
      Caption         =   "Referencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   180
      TabIndex        =   104
      Top             =   5475
      Width           =   3570
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         Caption         =   "Mesa Abierta"
         Height          =   240
         Left            =   150
         TabIndex        =   107
         Top             =   195
         Width           =   1020
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   "Mesa Pendiente"
         Height          =   240
         Left            =   1215
         TabIndex        =   106
         Top             =   195
         Width           =   1170
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C000&
         Caption         =   "Mesa Vacía"
         Height          =   240
         Left            =   2475
         TabIndex        =   105
         Top             =   210
         Width           =   960
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   360
      Left            =   4020
      TabIndex        =   51
      Top             =   5655
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   4830
      Left            =   90
      TabIndex        =   1
      Top             =   525
      Width           =   4905
      Begin VB.Frame FrmMesas 
         Caption         =   "Mesas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4500
         Left            =   105
         TabIndex        =   2
         Top             =   195
         Width           =   4650
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Left            =   1965
            Top             =   4410
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   50
            Left            =   4170
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   1902
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   49
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   1902
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   48
            Left            =   3270
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   1902
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   47
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   1902
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   46
            Left            =   2370
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   1902
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   45
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   1902
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1"
            Height          =   345
            Index           =   44
            Left            =   1470
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   1902
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   43
            Left            =   1020
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   1902
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "10"
            Height          =   345
            Index           =   42
            Left            =   570
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   1902
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "10"
            Height          =   345
            Index           =   41
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   1902
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   60
            Left            =   4170
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   2310
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   59
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   2310
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   58
            Left            =   3270
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   2310
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   57
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   2310
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   56
            Left            =   2370
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   2310
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   55
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   2310
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1"
            Height          =   345
            Index           =   54
            Left            =   1470
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   2310
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   53
            Left            =   1020
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   2310
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "10"
            Height          =   345
            Index           =   52
            Left            =   570
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   2310
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "10"
            Height          =   345
            Index           =   51
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   2310
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   100
            Left            =   4170
            Style           =   1  'Graphical
            TabIndex        =   103
            Top             =   3945
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   99
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   102
            Top             =   3945
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   98
            Left            =   3270
            Style           =   1  'Graphical
            TabIndex        =   101
            Top             =   3945
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   97
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   100
            Top             =   3945
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   96
            Left            =   2370
            Style           =   1  'Graphical
            TabIndex        =   99
            Top             =   3945
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   95
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   98
            Top             =   3945
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1"
            Height          =   345
            Index           =   94
            Left            =   1470
            Style           =   1  'Graphical
            TabIndex        =   97
            Top             =   3945
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   93
            Left            =   1020
            Style           =   1  'Graphical
            TabIndex        =   96
            Top             =   3945
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "10"
            Height          =   345
            Index           =   92
            Left            =   570
            Style           =   1  'Graphical
            TabIndex        =   95
            Top             =   3945
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "10"
            Height          =   345
            Index           =   91
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   94
            Top             =   3945
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   90
            Left            =   4170
            Style           =   1  'Graphical
            TabIndex        =   93
            Top             =   3534
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   89
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   92
            Top             =   3534
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   88
            Left            =   3270
            Style           =   1  'Graphical
            TabIndex        =   91
            Top             =   3534
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   87
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   90
            Top             =   3534
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   86
            Left            =   2370
            Style           =   1  'Graphical
            TabIndex        =   89
            Top             =   3534
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   85
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   88
            Top             =   3534
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1"
            Height          =   345
            Index           =   84
            Left            =   1470
            Style           =   1  'Graphical
            TabIndex        =   87
            Top             =   3534
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   83
            Left            =   1020
            Style           =   1  'Graphical
            TabIndex        =   86
            Top             =   3534
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "10"
            Height          =   345
            Index           =   82
            Left            =   570
            Style           =   1  'Graphical
            TabIndex        =   85
            Top             =   3534
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "10"
            Height          =   345
            Index           =   81
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   84
            Top             =   3534
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   70
            Left            =   4170
            Style           =   1  'Graphical
            TabIndex        =   73
            Top             =   2718
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   69
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   72
            Top             =   2718
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   68
            Left            =   3270
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   2718
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   67
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   2718
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   66
            Left            =   2370
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   2718
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   65
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   2718
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1"
            Height          =   345
            Index           =   64
            Left            =   1470
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   2718
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   63
            Left            =   1020
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   2718
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "10"
            Height          =   345
            Index           =   62
            Left            =   570
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   2718
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "10"
            Height          =   345
            Index           =   61
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   2718
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   80
            Left            =   4170
            Style           =   1  'Graphical
            TabIndex        =   83
            Top             =   3126
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   79
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   82
            Top             =   3126
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   78
            Left            =   3270
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   3126
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   77
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   3126
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   76
            Left            =   2370
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   3126
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   75
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   78
            Top             =   3126
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1"
            Height          =   345
            Index           =   74
            Left            =   1470
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   3126
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   73
            Left            =   1020
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   3126
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "10"
            Height          =   345
            Index           =   72
            Left            =   570
            Style           =   1  'Graphical
            TabIndex        =   75
            Top             =   3126
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "10"
            Height          =   345
            Index           =   71
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   3126
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "10"
            Height          =   345
            Index           =   31
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   1494
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "10"
            Height          =   345
            Index           =   32
            Left            =   570
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   1494
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   33
            Left            =   1020
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   1494
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1"
            Height          =   345
            Index           =   34
            Left            =   1470
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   1494
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   35
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   1494
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   36
            Left            =   2370
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   1494
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   37
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   1494
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   38
            Left            =   3270
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   1494
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   39
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   1494
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   40
            Left            =   4170
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   1494
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   30
            Left            =   4170
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   1086
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   29
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   1086
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   28
            Left            =   3270
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   1086
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   27
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   1086
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   26
            Left            =   2370
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   1086
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   25
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1086
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1"
            Height          =   345
            Index           =   24
            Left            =   1470
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   1086
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1"
            Height          =   345
            Index           =   23
            Left            =   1020
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   1086
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "10"
            Height          =   345
            Index           =   22
            Left            =   570
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   1086
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "10"
            Height          =   345
            Index           =   21
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   1086
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   20
            Left            =   4170
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   678
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   19
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   678
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   18
            Left            =   3270
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   678
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   17
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   678
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   16
            Left            =   2370
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   678
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   15
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   678
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   14
            Left            =   1470
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   678
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   13
            Left            =   1020
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   678
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "10"
            Height          =   345
            Index           =   12
            Left            =   570
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   678
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "10"
            Height          =   345
            Index           =   11
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   678
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            BackColor       =   &H00C0C0C0&
            Caption         =   "10"
            Height          =   345
            Index           =   10
            Left            =   4170
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   270
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "10"
            Height          =   345
            Index           =   9
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   270
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   8
            Left            =   3270
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   270
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   7
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   270
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   6
            Left            =   2370
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   270
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   5
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   270
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   4
            Left            =   1470
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   270
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            Caption         =   "1"
            Height          =   345
            Index           =   3
            Left            =   1020
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   270
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1"
            Height          =   345
            Index           =   2
            Left            =   570
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   270
            Width           =   390
         End
         Begin VB.CommandButton CmdMesa 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1"
            Height          =   345
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   270
            Width           =   390
         End
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Mesa :"
      Height          =   210
      Left            =   180
      TabIndex        =   108
      Top             =   105
      Width           =   570
   End
End
Attribute VB_Name = "FrmAdicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim Handler As Integer

Public Sub Refrescar(nMesa As Byte)
  Dim Rx As Recordset, cRsl As ClsLectura, sCondi As String
  
  Set cRsl = New ClsLectura
  
  sCondi = "A_Mesa=" & nMesa & " and a_mensaje<>'DELIVERY'"
'
  Set Rx = cRsl.TraerRsCondi("Auxiliar", "A_Mensaje", sCondi)
      
  If Rx.RecordCount <> 0 Then
     Select Case Rx![A_MENSAJE]
            Case "ABIERTA"
                 CmdMesa(nMesa).BackColor = &HFF&
            Case "PENDIENTE"
                 CmdMesa(nMesa).BackColor = &HFF8080
            Case "VACIA"
                 CmdMesa(nMesa).BackColor = &HC000&
            Case "A RENDIR"
                 CmdMesa(nMesa).BackColor = &HC000&
     End Select
  Else
     CmdMesa(nMesa).BackColor = &HC000&
  End If
  
  Set cRsl = Nothing
  Rx.Close
  
End Sub

Private Sub CmdMesa_Click(Index As Integer)
  Dim ok As Boolean
  
  Dim Mesas As ClsPrograma
  Set Mesas = New ClsPrograma
  
  nMesa = Index
  
  ok = Mesas.DameEstadoMesa("ABIERTA", nMesa)
  bDelivery = False
  
  If ok Then
     bAbri = False
     frmMesa.Show 1
     If TxtMesa.Enabled = True Then
        TxtMesa.SetFocus
     End If
  Else
     MenuMesa.MenuAbrir_Click
  End If

  Set Mesas = Nothing
  
End Sub

Private Sub CmdMesa_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF1 Then
     nMesa = Index
     Select Case FrmAdicion.CmdMesa(nMesa).BackColor
            Case 49152  ' Verde Vacia
                 MenuMesa.mnuabrir.Enabled = True
                 MenuMesa.MnuCopiar.Enabled = False
                 MenuMesa.CierreRapido.Enabled = False
                 MenuMesa.MnuPendiente.Enabled = False
                 MenuMesa.mncobranza.Enabled = False
                 MenuMesa.MenuAnular.Enabled = False
            Case 255  'Rojo abierta
                 MenuMesa.mnuabrir.Enabled = False
                 MenuMesa.MnuCopiar.Enabled = True
                 MenuMesa.CierreRapido.Enabled = True
                 MenuMesa.MnuPendiente.Enabled = True
                 MenuMesa.mncobranza.Enabled = False
                 MenuMesa.MenuAnular.Enabled = True
            Case 16744576
                 MenuMesa.mnuabrir.Enabled = False
                 MenuMesa.MnuCopiar.Enabled = False
                 MenuMesa.CierreRapido.Enabled = False
                 MenuMesa.MnuPendiente.Enabled = False
                 MenuMesa.mncobranza.Enabled = True
                 MenuMesa.MenuAnular.Enabled = False
     End Select
     FrmAdicion.PopupMenu MenuMesa.mnumesa
   End If
End Sub

Private Sub CmdMesa_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
     nMesa = Index
     Select Case FrmAdicion.CmdMesa(nMesa).BackColor
            Case 49152  ' Verde Vacia
                 MenuMesa.mnuabrir.Enabled = True
                 MenuMesa.MnuCopiar.Enabled = False
                 MenuMesa.CierreRapido.Enabled = False
                 MenuMesa.MnuPendiente.Enabled = False
                 MenuMesa.mncobranza.Enabled = False
                 MenuMesa.MenuAnular.Enabled = False
            Case 255  'Rojo abierta
                 MenuMesa.mnuabrir.Enabled = False
                 MenuMesa.MnuCopiar.Enabled = True
                 MenuMesa.CierreRapido.Enabled = True
                 MenuMesa.MnuPendiente.Enabled = True
                 MenuMesa.mncobranza.Enabled = False
                 MenuMesa.MenuAnular.Enabled = True
            Case 16744576
                 MenuMesa.mnuabrir.Enabled = False
                 MenuMesa.MnuCopiar.Enabled = False
                 MenuMesa.CierreRapido.Enabled = False
                 MenuMesa.MnuPendiente.Enabled = False
                 MenuMesa.mncobranza.Enabled = True
                 MenuMesa.MenuAnular.Enabled = False
     End Select
     FrmAdicion.PopupMenu MenuMesa.mnumesa
  End If

End Sub

'Private Sub CmdMesa_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Dim cRsl As ClsLectura
'
'   Set cRsl = New ClsLectura
'
'    CmdMesa(Index).ToolTipText = cRsl.TraerValorDeUnCampo("Clientes", "RazonSocial", "Id=" & cRsl.TraerValorDeUnCampo("Auxiliar", "A_CLIENTE", "A_MESA=" & CmdMesa(Index).Caption & " AND A_CODART='01-0099'"))
'
'    Set cRsl = Nothing
'
'End Sub

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Dim i As Byte
  TxtMesa.text = ""
  For i = 1 To 100
      CmdMesa(i).Caption = i
      Refrescar i
  Next
  Me.Top = 0
  Me.Left = 0
End Sub

