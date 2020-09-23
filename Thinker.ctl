VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Thinker 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H00C0E0FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin Project1.ExtendedTimer ExtendedTimer1 
      Left            =   1755
      Top             =   2880
      _extentx        =   741
      _extenty        =   741
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1545
      Left            =   960
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   235
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   3555
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2640
      Top             =   2625
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   16711935
      _Version        =   393216
   End
   Begin Project1.Mask Mask1 
      Height          =   210
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   4455
      _extentx        =   7858
      _extenty        =   370
      hold_barspace   =   8
   End
End
Attribute VB_Name = "Thinker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const Border_Top_Left = "2,16711935,16711935,6842472,16711935,7961211,15724527,6842472,12500670,15724527"
Private Const Border_Top_Right = "2,6842472,12500670,15724527,16711935,7961211,15724527,16711935,16711935,6842472"
Private Const Border_Bottom_Left = "2,6842472,16711935,16711935,15724527,7961211,16711935,-1,15724527,6842472"
Private Const Border_Bottom_Right = "2,-1,15724527,6842472,15724527,7961211,16711935,6842472,16711935,16711935"
Private Const BlankBar = "22,-1,-1,-1,150,100,60,40,30,20,10,1,1,1,1,1,1,1,10,20,40,60,100,150"
Private Const BlankBar_Solid = "22,-1,-1,-1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1"
Private Const Square_Reg = "10,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,16777215,16777215,16777215,16777215,16777215,16777215,16777215,0,0,0,0,16777215,16777215,16777215,16777215,16777215,16777215,16777215,0,0,0,0,16777215,16777215,16777215,16777215,16777215,16777215,16777215,0,0,0,0,16777215,16777215,16777215,16777215,16777215,16777215,16777215,0,0,0,0,16777215,16777215,16777215,16777215,16777215,16777215,16777215,0,0,0,0,16777215,16777215,16777215,16777215,16777215,16777215,16777215,0,0,0,0,16777215,16777215,16777215,16777215,16777215,16777215,16777215,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
Private Const Blank_Square = "10,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,0,0,0,0,0,0,0,-1,-1,-1,-1,0,0,0,0,0,0,0,-1,-1,-1,-1,0,0,0,0,0,0,0,-1,-1,-1,-1,0,0,0,0,0,0,0,-1,-1,-1,-1,0,0,0,0,0,0,0,-1,-1,-1,-1,0,0,0,0,0,0,0,-1,-1,-1,-1,0,0,0,0,0,0,0,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1"
Private Const Square_Sides = "8,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,0,0,0,0,0,0,0,-1,-1,0,0,0,0,0,0,0,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,0,0,0,0,0,0,0,-1,-1,0,0,0,0,0,0,0"
Private Const Blank_Square2 = "10,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1"
Private Const Square_Reg2 = "10,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215," & _
"16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215"

'Dim WithEvents s As clsTimer

Public Enum BorderStyle_
    [Round_Bar]
    [Square_Bar]
    [Square_Spacers]
End Enum

Private Hold_BorderStyle As BorderStyle_
Private Hold_BackColor As OLE_COLOR
Private Hold_BarColor As OLE_COLOR
Private Hold_BarSpace As Integer
Private Hold_BarWidth As Integer
Private Hold_Animate As Boolean
Private Hold_BarCount As Integer
Private Hold_Fade As Integer
Private Hold_Speed As Long
Private Hold_BarBorder As Boolean
Private Hold_Bounce As Boolean

Private ThinkerPlace As Integer
Private BarsPer As Integer
Private Positive As Boolean

Property Let Bounce(StrValue As Boolean)
    Hold_Bounce = StrValue
End Property

Property Get Bounce() As Boolean
    BarBorder = Hold_Bounce
End Property

Property Let BarBorder(StrValue As Boolean)
    Hold_BarBorder = StrValue
    If Hold_BorderStyle = Square_Spacers Then Mask1.Visible = StrValue
End Property

Property Get BarBorder() As Boolean
    BarBorder = Hold_BarBorder
End Property

Property Let Speed(StrValue As Long)
    Hold_Speed = StrValue
End Property

Property Get Speed() As Long
    Speed = Hold_Speed
End Property
Property Let Fade(StrValue As Integer)
    Hold_Fade = StrValue
    LoadGUI
End Property

Property Get Fade() As Integer
    Fade = Hold_Fade
End Property

Property Let BarCount(StrValue As Integer)
    Hold_BarCount = StrValue
    LoadGUI
End Property

Property Get BarCount() As Integer
    BarCount = Hold_BarCount
End Property

Property Let Animate(StrValue As Boolean)
    Hold_Animate = StrValue
    Select Case Hold_Animate
    Case True
        ExtendedTimer1.Interval = Hold_Speed
        ExtendedTimer1.Enabled = True
    Case False
        ExtendedTimer1.Enabled = False
    End Select
End Property

Property Get Animate() As Boolean
    Animate = Hold_Animate
End Property

Property Let BarWidth(StrValue As Integer)
    Hold_BarWidth = StrValue
    LoadGUI
End Property

Property Get BarWidth() As Integer
    BarWidth = Hold_BarWidth
End Property

Property Let BarSpace(StrValue As Integer)
    Hold_BarSpace = StrValue
    LoadGUI
End Property

Property Get BarSpace() As Integer
    BarSpace = Hold_BarSpace
End Property

Property Let BackColor(StrValue As OLE_COLOR)
    Hold_BackColor = StrValue
    LoadGUI
End Property

Property Get BackColor() As OLE_COLOR
    BackColor = Hold_BackColor
End Property

Property Let BarColor(StrValue As OLE_COLOR)
    Hold_BarColor = StrValue
    LoadGUI
End Property

Property Get BarColor() As OLE_COLOR
    BarColor = Hold_BarColor
End Property

Property Let Style(StrValue As BorderStyle_)
    Hold_BorderStyle = StrValue
    LoadGUI
End Property

Property Get Style() As BorderStyle_
    Style = Hold_BorderStyle
End Property
Private Function LoadClearBmp(Brightness As Integer, AddClor As Long, Legnth As Integer, ColorPallet As String, X As Integer, Y As Integer) As Integer
    If ColorPallet = "" Then Exit Function
    Dim PixCount
    Dim Colors() As String, CurrentRow, CurrentColumn, Count, Rows
    Colors = Split(ColorPallet, ",")
    Rows = Int(Split(ColorPallet, ",")(0))
    For Count = 1 To UBound(Colors)
        If CurrentRow > (Rows) Then CurrentRow = 0: CurrentColumn = CurrentColumn + 1
            If Colors(Count) <> -1 Then
                    UserControl.Line (X + CurrentColumn, Y + CurrentRow)-(X + CurrentColumn + Legnth, Y + CurrentRow), AdjustBrightness(AddClor, Colors(Count) + Brightness)
            End If
        CurrentRow = CurrentRow + 1
    Next
    LoadClearBmp = CurrentColumn
End Function
Private Function LoadBmpMenuLines(Legnth As Integer, ColorPallet As String, X As Integer, Y As Integer, Optional Gray As Boolean = True, Optional Brightness As Integer) As Integer
    If ColorPallet = "" Then Exit Function
    Dim PixCount
    Dim Colors() As String, CurrentRow, CurrentColumn, Count, Rows
    Colors = Split(ColorPallet, ",")
    Rows = Int(Split(ColorPallet, ",")(0))
    For Count = 1 To UBound(Colors)
        If CurrentRow > (Rows) Then CurrentRow = 0: CurrentColumn = CurrentColumn + 1
            If Colors(Count) <> -1 Then
                    Picture1.Line (X + CurrentColumn, Y + CurrentRow)-(X + CurrentColumn + Legnth, Y + CurrentRow), MakeGrey(Colors(Count))
            End If
        CurrentRow = CurrentRow + 1
    Next
    LoadBmpMenuLines = CurrentColumn
End Function

Private Function LoadBmp(Legnth As Integer, ColorPallet As String, X As Integer, Y As Integer, Optional Gray As Boolean = True, Optional Brightness As Integer, Optional IgnoreColor As OLE_COLOR = vbBlack) As Integer
    If ColorPallet = "" Then Exit Function
    Dim PixCount
    Dim Colors() As String, CurrentRow, CurrentColumn, Count, Rows
    Colors = Split(ColorPallet, ",")
    Rows = Int(Split(ColorPallet, ",")(0))
    For Count = 1 To UBound(Colors)
        If CurrentRow > (Rows) Then CurrentRow = 0: CurrentColumn = CurrentColumn + 1
            If Colors(Count) <> -1 And IgnoreColor <> Colors(Count) Then
                UserControl.Line (X + CurrentColumn, Y + CurrentRow)-(X + CurrentColumn + Legnth, Y + CurrentRow), MakeGrey(Colors(Count))
            End If
        CurrentRow = CurrentRow + 1
    Next
    LoadBmp = CurrentColumn
End Function
Function LoadGUI()
DrawBorder
LoadShell
ThinkerPlace = -6
If Hold_BorderStyle = Square_Spacers Then
BarsPer = (UserControl.ScaleWidth / (3 + (Hold_BarWidth))) / 3
Else
BarsPer = UserControl.ScaleWidth / (3 + (Hold_BarWidth + Hold_BarSpace))
End If
Positive = True
If Hold_BorderStyle = Square_Spacers Then Mask1.Visible = Hold_BarBorder
End Function

Function LoadBars(Start As Integer, BarCount As Integer)
    Dim X As Integer, Y, Z As Integer
If Hold_BorderStyle = Square_Spacers Then
    X = 3 + (Start * (Hold_BarWidth + Hold_BarSpace))
    For Y = 1 To BarCount
        If Positive = True Then
            LoadClearBmp (BarCount - Y) * Hold_Fade, Hold_BarColor, Hold_BarWidth, Blank_Square2, X, 1
        Else
            LoadClearBmp Y * Hold_Fade, Hold_BarColor, Hold_BarWidth, Blank_Square2, X, 1
        End If
        X = X + Hold_BarWidth
    Next
Else
    X = 3 + (Start * (Hold_BarWidth + Hold_BarSpace))
    For Y = 1 To BarCount
        If Positive = True Then
        LoadClearBmp (BarCount - Y) * Hold_Fade, Hold_BarColor, Hold_BarWidth, BlankBar_Solid, X, 0
        Else
        LoadClearBmp Y * Hold_Fade, Hold_BarColor, Hold_BarWidth, BlankBar_Solid, X, 0
        End If
        X = X + Hold_BarWidth + Hold_BarSpace
    Next
End If
End Function

Function LoadSquares(BarCount As Integer)
    Dim X As Integer, Y
    X = 3
    For Y = 1 To BarCount
        X = X + LoadBmpMenuLines(1, Square_Reg2, X + 1, 1) + Hold_BarSpace + 2
    Next
End Function
Function LoadSquares2(BarCount As Integer)
    Dim X As Integer, Y
    X = 3
    For Y = 1 To BarCount
        X = X + LoadBmp(1, Square_Sides, X, 1, , , vbWhite) + Hold_BarSpace + 1
    Next
End Function

Function ClearBar(Place As Integer)
If Hold_BorderStyle = Square_Spacers Then
    Place = 3 + (Place * (Hold_BarWidth + 6))
    LoadClearBmp 0, Hold_BackColor, Hold_BarWidth + 6, Blank_Square2, Place, 1
Else
    Place = 3 + (Place * (Hold_BarWidth + Hold_BarSpace))
    LoadClearBmp 0, Hold_BackColor, Hold_BarWidth, BlankBar_Solid, Place, 0
End If
End Function

Function DrawBorder()
    Picture1.Cls
    Select Case Hold_BorderStyle
    Case [Round_Bar]
        Picture1.BackColor = Hold_BackColor
        Picture1.Line (0, 0)-(UserControl.ScaleWidth, 0), 6842472
        Picture1.Line (0, 1)-(UserControl.ScaleWidth, 1), 12500670
        Picture1.Line (0, 2)-(UserControl.ScaleWidth, 2), 15724527
        
        Picture1.Line (0, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth, UserControl.ScaleHeight - 1), 6842472
        Picture1.Line (0, UserControl.ScaleHeight - 2)-(UserControl.ScaleWidth, UserControl.ScaleHeight - 2), 15724527
        
        Picture1.Line (0, 0)-(0, UserControl.ScaleHeight), 6842472
        Picture1.Line (1, 0)-(1, UserControl.ScaleHeight), 15724527
        
        Picture1.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), 6842472
        Picture1.Line (UserControl.ScaleWidth - 2, 0)-(UserControl.ScaleWidth - 2, UserControl.ScaleHeight), 15724527
        
        LoadBmpMenuLines 1, Border_Top_Left, 0, 0
        LoadBmpMenuLines 1, Border_Top_Right, UserControl.ScaleWidth - 3, 0
        LoadBmpMenuLines 1, Border_Bottom_Left, 0, UserControl.ScaleHeight - 3
        LoadBmpMenuLines 1, Border_Bottom_Right, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3
    Case [Square_Bar]
        Picture1.BackColor = Hold_BackColor
        Picture1.Line (0, 0)-(UserControl.ScaleWidth - 1, 0), 10070188
        Picture1.Line (1, 1)-(UserControl.ScaleWidth - 2, 1), 6582129
        Picture1.Line (0, 0)-(0, UserControl.ScaleHeight - 1), 10070188
        Picture1.Line (1, 1)-(1, UserControl.ScaleHeight - 2), 6582129
        
        Picture1.Line (1, UserControl.ScaleHeight - 2)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 2), 14872561
        Picture1.Line (UserControl.ScaleWidth - 2, 1)-(UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 1), 14872561
        Picture1.Line (1, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), vbWhite
        Picture1.Line (UserControl.ScaleWidth - 1, 1)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), vbWhite
    Case [Square_Spacers]
        Picture1.BackColor = &HFF00FF
        LoadSquares Hold_BarCount
    End Select
End Function

Function DrawSide(Side As Integer)
    Select Case Hold_BorderStyle
    Case [Round_Bar]
        Select Case Side
        Case 0
        UserControl.Line (0, 0)-(0, UserControl.ScaleHeight), 6842472
        UserControl.Line (1, 0)-(1, UserControl.ScaleHeight), 15724527
        LoadBmp 1, Border_Top_Left, 0, 0
        LoadBmp 1, Border_Bottom_Left, 0, UserControl.ScaleHeight - 3
        Case 1
        UserControl.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), 6842472
        UserControl.Line (UserControl.ScaleWidth - 2, 0)-(UserControl.ScaleWidth - 2, UserControl.ScaleHeight), 15724527
        LoadBmp 1, Border_Top_Right, UserControl.ScaleWidth - 3, 0
        LoadBmp 1, Border_Bottom_Right, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3
        End Select
    Case [Square_Bar]
        'Select Case Side
        'Case 0
        Picture1.Line (0, 0)-(0, UserControl.ScaleHeight - 1), 10070188
        Picture1.Line (1, 1)-(1, UserControl.ScaleHeight - 2), 6582129
        'Case 1
        Picture1.Line (1, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), vbWhite
        Picture1.Line (UserControl.ScaleWidth - 1, 1)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), vbWhite
        'End Select
    End Select
End Function

Function LoadShell()
    ImageList1.ListImages.Add 1, , Picture1.Image
    UserControl.Picture = ImageList1.ListImages(1).Picture
    UserControl.MaskPicture = UserControl.Picture
    ImageList1.ListImages.Remove 1
End Function

Private Sub ExtendedTimer1_Timer()
    If Positive = True Then
        ClearBar ThinkerPlace - 1
    Else
        If Hold_BorderStyle = Square_Spacers Then
            ClearBar ThinkerPlace + Hold_BarCount / 3
        Else
            ClearBar ThinkerPlace + Hold_BarCount
        End If
    End If
    
    If Hold_BorderStyle = Square_Spacers Then
        LoadBars ThinkerPlace, Hold_BarCount / 3
    Else
        LoadBars ThinkerPlace, Hold_BarCount
    End If
    
    If Positive = True Then
        ThinkerPlace = ThinkerPlace + 1
        DrawSide 0
    Else
        ThinkerPlace = ThinkerPlace - 1
        DrawSide 1
    End If
    
    If ThinkerPlace - (Hold_BarCount + 4) > BarsPer Then
        If Hold_Bounce = True Then
            Positive = False
        Else
            Positive = True
            If Hold_BorderStyle = Square_Spacers Then
                ClearBar ThinkerPlace + Hold_BarCount / 3
            Else
                ClearBar ThinkerPlace + Hold_BarCount
            End If
            ThinkerPlace = 0
            DrawSide 1
        End If
    ElseIf ThinkerPlace < -(Hold_BarCount + 2) Then
        Positive = True
    End If
    DoEvents
End Sub

Private Sub UserControl_Resize()
    Picture1.Top = 0
    Picture1.Left = 0
    Picture1.Width = UserControl.Width / 14
    Picture1.Height = UserControl.Height / 14
    LoadGUI
End Sub

Private Sub UserControl_Show()
    LoadGUI
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Hold_BorderStyle = PropBag.ReadProperty("Hold_BorderStyle", [Round_Bar])
    Hold_BackColor = PropBag.ReadProperty("Hold_BackColor", vbWhite)
    Hold_BarColor = PropBag.ReadProperty("Hold_BarColor", &HC000&)
    Hold_BarSpace = PropBag.ReadProperty("Hold_BarSpace", 2)
    Hold_BarWidth = PropBag.ReadProperty("Hold_BarWidth", 6)
    Hold_BarCount = PropBag.ReadProperty("Hold_BarCount", 6)
    Hold_Fade = PropBag.ReadProperty("Hold_Fade", 40)
    Hold_Speed = PropBag.ReadProperty("Hold_Speed", 50)
    Hold_BarBorder = PropBag.ReadProperty("Hold_BarBorder", True)
    Hold_Bounce = PropBag.ReadProperty("Hold_Bounce", True)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Hold_BorderStyle", Hold_BorderStyle, [Round_Bar]
    PropBag.WriteProperty "Hold_BackColor", Hold_BackColor, vbWhite
    PropBag.WriteProperty "Hold_BarColor", Hold_BarColor, &HC000&
    PropBag.WriteProperty "Hold_BarSpace", Hold_BarSpace, 2
    PropBag.WriteProperty "Hold_BarWidth", Hold_BarWidth, 6
    PropBag.WriteProperty "Hold_BarCount", Hold_BarCount, 6
    PropBag.WriteProperty "Hold_Fade", Hold_Fade, 40
    PropBag.WriteProperty "Hold_Speed", Hold_Speed, 50
    PropBag.WriteProperty "Hold_BarBorder", Hold_BarBorder, True
    PropBag.WriteProperty "Hold_Bounce", Hold_Bounce, True
End Sub
