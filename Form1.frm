VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6930
   ClientLeft      =   1890
   ClientTop       =   -2400
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00400000&
      Height          =   825
      Left            =   -45
      ScaleHeight     =   765
      ScaleWidth      =   7695
      TabIndex        =   32
      Top             =   -30
      Width           =   7755
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6180
         TabIndex        =   34
         Top             =   390
         Width           =   1830
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Waiting Status Bars"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   390
         TabIndex        =   33
         Top             =   150
         Width           =   5130
      End
   End
   Begin VB.PictureBox Hold_p 
      BackColor       =   &H00FFFFFF&
      Height          =   5085
      Left            =   4935
      ScaleHeight     =   5025
      ScaleWidth      =   2640
      TabIndex        =   16
      Top             =   930
      Width           =   2700
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   900
         ScaleHeight     =   150
         ScaleWidth      =   150
         TabIndex        =   18
         Top             =   690
         Width           =   180
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   900
         ScaleHeight     =   150
         ScaleWidth      =   150
         TabIndex        =   17
         Top             =   900
         Width           =   180
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "BarBorder: "
         Height          =   255
         Left            =   45
         TabIndex        =   31
         Top             =   1725
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Currently Viewing Sample: "
         Height          =   285
         Left            =   60
         TabIndex        =   30
         Top             =   105
         Width           =   2880
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Index: "
         Height          =   240
         Left            =   60
         TabIndex        =   29
         Top             =   300
         Width           =   2040
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Style: "
         Height          =   240
         Left            =   60
         TabIndex        =   28
         Top             =   495
         Width           =   2715
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "BackColor:"
         Height          =   180
         Left            =   45
         TabIndex        =   27
         Top             =   690
         Width           =   1050
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         Height          =   180
         Left            =   1080
         TabIndex        =   26
         Top             =   675
         Width           =   1200
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "BarColor:"
         Height          =   210
         Left            =   45
         TabIndex        =   25
         Top             =   915
         Width           =   780
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         Height          =   180
         Left            =   1080
         TabIndex        =   24
         Top             =   900
         Width           =   1200
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "BarCount: "
         Height          =   270
         Left            =   45
         TabIndex        =   23
         Top             =   1125
         Width           =   930
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "BarSpace:"
         Height          =   180
         Left            =   45
         TabIndex        =   22
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "BarWidth:"
         Height          =   270
         Left            =   45
         TabIndex        =   21
         Top             =   1530
         Width           =   1230
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Fade:"
         Height          =   270
         Left            =   45
         TabIndex        =   20
         Top             =   1905
         Width           =   1230
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Speed: "
         Height          =   315
         Left            =   45
         TabIndex        =   19
         Top             =   2085
         Width           =   1605
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Sample4"
      Height          =   465
      Index           =   3
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6375
      Width           =   1425
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Sample2"
      Height          =   465
      Index           =   2
      Left            =   2220
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6345
      Width           =   1425
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Sample3"
      Height          =   465
      Index           =   1
      Left            =   4020
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6360
      Width           =   1425
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Sample1"
      Height          =   465
      Index           =   0
      Left            =   390
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6345
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sample 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   3
      Left            =   105
      TabIndex        =   9
      Top             =   4785
      Width           =   4785
      Begin Project1.Thinker Thinker1 
         Height          =   375
         Index           =   3
         Left            =   75
         TabIndex        =   10
         Top             =   765
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   661
         Hold_BorderStyle=   1
         Hold_BackColor  =   -2147483633
         Hold_BarColor   =   16576
         Hold_Fade       =   45
         Hold_Speed      =   25
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Building drivers list, please wait..."
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   510
         Width           =   4560
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sample 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   2
      Left            =   75
      TabIndex        =   6
      Top             =   2160
      Width           =   4785
      Begin Project1.Thinker Thinker1 
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   765
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   423
         Hold_BorderStyle=   2
         Hold_BarColor   =   192
         Hold_BarSpace   =   8
         Hold_BarWidth   =   10
         Hold_BarCount   =   15
         Hold_Fade       =   45
         Hold_Speed      =   25
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Searching for availible updates, please wait..."
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   510
         Width           =   4560
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sample 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   75
      TabIndex        =   3
      Top             =   3465
      Width           =   4785
      Begin Project1.Thinker Thinker1 
         Height          =   375
         Index           =   1
         Left            =   75
         TabIndex        =   4
         Top             =   765
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   661
         Hold_BarColor   =   32768
         Hold_Fade       =   45
         Hold_Speed      =   25
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Checking for program version..."
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   510
         Width           =   4560
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sample 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   840
      Width           =   4785
      Begin Project1.Thinker Thinker1 
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   765
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   423
         Hold_BorderStyle=   2
         Hold_BarColor   =   8388608
         Hold_BarSpace   =   8
         Hold_BarWidth   =   10
         Hold_BarCount   =   15
         Hold_Fade       =   45
         Hold_Speed      =   25
         Hold_BarBorder  =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Please Wait while windows searches for installed componets..."
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   510
         Width           =   4560
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      X1              =   90
      X2              =   7500
      Y1              =   6180
      Y2              =   6180
   End
   Begin VB.Line Line1 
      DrawMode        =   1  'Blackness
      X1              =   75
      X2              =   7485
      Y1              =   6165
      Y2              =   6165
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Chosen As Integer

Private Sub Form_Load()
Chosen = -1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Chosen <> -1 Then Thinker1(Chosen).Animate = False
DoEvents
End Sub

Private Sub Option1_Click(Index As Integer)
Label2.Caption = "Currently Viewing Sample: " & Index + 1
Label3.Caption = "Index: " & Index
Label4.Caption = "Style: " & Thinker1(Index).Style
Label6.Caption = Thinker1(Index).BackColor
Picture1.BackColor = Thinker1(Index).BackColor
Label8.Caption = Thinker1(Index).BarColor
Picture2.BackColor = Thinker1(Index).BarColor
Label9.Caption = "BarCount: " & Thinker1(Index).BarCount
Label10.Caption = "BarSpace: " & Thinker1(Index).BarSpace
Label11.Caption = "BarWidth: " & Thinker1(Index).BarWidth
Label12.Caption = "Fade: " & Thinker1(Index).Fade
Label14.Caption = "Speed: " & Thinker1(Index).Speed
Label13.Caption = "BarBorder: " & Thinker1(Index).BarBorder
Thinker1(Index).Animate = True
If Chosen <> -1 Then Thinker1(Chosen).Animate = False
Chosen = Index
End Sub
