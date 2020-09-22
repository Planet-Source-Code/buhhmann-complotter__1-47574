VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  '2D
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "ComPlotter"
   ClientHeight    =   4530
   ClientLeft      =   1575
   ClientTop       =   1485
   ClientWidth     =   5070
   FillColor       =   &H00404040&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "IOtest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   Picture         =   "IOtest.frx":0442
   ScaleHeight     =   302
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   Begin VB.CheckBox Check9 
      Caption         =   "Repaint"
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLS"
      Height          =   495
      Left            =   1320
      TabIndex        =   13
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3480
      TabIndex        =   10
      Text            =   "0,1"
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Starten + Stoppen"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "Inputs"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      TabIndex        =   5
      Top             =   1680
      Width           =   5055
      Begin VB.CheckBox Check5 
         Appearance      =   0  '2D
         BackColor       =   &H00FF00FF&
         Caption         =   "RXD"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4080
         TabIndex        =   17
         ToolTipText     =   "ACHTUNG! RXD ist anders! (-1 für positiv, 0 für negativ)"
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  '2D
         BackColor       =   &H000000FF&
         Caption         =   "CTS"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  '2D
         BackColor       =   &H0000FF00&
         Caption         =   "DSR"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1200
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  '2D
         BackColor       =   &H00FFFF00&
         Caption         =   "DCD"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2160
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox Check4 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         Caption         =   "RI"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   3120
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Outputs"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   5055
      Begin VB.CheckBox Check8 
         BackColor       =   &H00800000&
         Caption         =   "RTS"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   3600
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00008000&
         Caption         =   "DTR"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00000080&
         Caption         =   "TXD"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   720
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      Caption         =   "Graph"
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   5055
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'Kein
         Height          =   1300
         Left            =   480
         ScaleHeight     =   87
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   300
         TabIndex        =   15
         Top             =   265
         Width           =   4500
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   88
         End
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "RTS"
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   0
         TabIndex        =   25
         Top             =   1388
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "DTR"
         ForeColor       =   &H00008000&
         Height          =   165
         Left            =   0
         TabIndex        =   24
         Top             =   1223
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "TXD"
         ForeColor       =   &H00000080&
         Height          =   165
         Left            =   0
         TabIndex        =   23
         Top             =   1058
         Width           =   495
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "RXD"
         ForeColor       =   &H00FF00FF&
         Height          =   165
         Left            =   0
         TabIndex        =   22
         Top             =   893
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "RI"
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Left            =   0
         TabIndex        =   21
         Top             =   728
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "DCD"
         ForeColor       =   &H00FFFF00&
         Height          =   165
         Left            =   0
         TabIndex        =   20
         Top             =   564
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "DSR"
         ForeColor       =   &H0000FF00&
         Height          =   165
         Left            =   0
         TabIndex        =   19
         Top             =   402
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "CTS"
         ForeColor       =   &H000000FF&
         Height          =   165
         Left            =   0
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Sec"
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Timerfrequenz:"
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   3960
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer
Dim CTSval As Boolean
Dim DSRval As Boolean
Dim DCDval As Boolean
Dim RIval As Boolean
Dim RXDval As Boolean
Dim TXDval As Boolean
Dim DTRval As Boolean
Dim RTSval As Boolean

Dim PLOT As Boolean

Private Sub Command1_Click()
If PLOT = True Then PLOT = False Else PLOT = True: Plotting
End Sub

Private Sub Command2_Click()
Picture1.Cls: X = 0
End Sub

Private Sub Form_Load()
    i = OPENCOM("COM2,1200,N,8,1")
 If i = 0 Then MsgBox ("COM Interface Error")
 TXD 0
 RTS 0
 DTR 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
  CLOSECOM
End Sub

Private Sub Plotting()
While PLOT = True: DoEvents
If Check6.Value Then TXD 1 Else TXD 0
If Check7.Value Then DTR 1 Else DTR 0
If Check8.Value Then RTS 1 Else RTS 0

CTSval = CTS()
DSRval = DSR()
DCDval = DCD()
RIval = RI()
RXDval = Not (READBYTE())
TXDval = Check6.Value
DTRval = Check7.Value
RTSval = Check8.Value

If CTSval = False Then Check1.Value = False
If CTSval = True Then Check1.Value = 1
If DSRval = False Then Check2.Value = False
If DSRval = True Then Check2.Value = 1
If DCDval = False Then Check3.Value = False
If DCDval = True Then Check3.Value = 1
If RIval = False Then Check4.Value = False
If RIval = True Then Check4.Value = 1
If RXDval = False Then Check5.Value = False
If RXDval = True Then Check5.Value = 1

X = X + 1
If X > Picture1.ScaleWidth Then X = 1: If Check9.Value = 1 Then Picture1.Cls

If CTSval = True Then Picture1.Line (X, 0)-(X, 10), Check1.BackColor
If DSRval = True Then Picture1.Line (X, 11)-(X, 21), Check2.BackColor
If DCDval = True Then Picture1.Line (X, 22)-(X, 32), Check3.BackColor
If RIval = True Then Picture1.Line (X, 33)-(X, 43), Check4.BackColor
If RXDval = True Then Picture1.Line (X, 44)-(X, 54), Check5.BackColor

If TXDval = True Then Picture1.Line (X, 55)-(X, 65), Check6.BackColor
If DTRval = True Then Picture1.Line (X, 66)-(X, 76), Check7.BackColor
If RTSval = True Then Picture1.Line (X, 77)-(X, 87), Check8.BackColor

Line1.X1 = X: Line1.X2 = X

Pause Text1.Text
Wend
End Sub

