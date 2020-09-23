VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Mask 
   BackColor       =   &H00000000&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   77
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   720
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   119
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3270
      Top             =   435
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "Mask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const Square_Reg = "10,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,16777215,16777215,16777215,16777215,16777215,16777215,16777215,0,0,0,0,16777215,16777215,16777215,16777215,16777215,16777215,16777215,0,0,0,0,16777215,16777215,16777215,16777215,16777215,16777215,16777215,0,0,0,0,16777215,16777215,16777215,16777215,16777215,16777215,16777215,0,0,0,0,16777215,16777215,16777215,16777215,16777215,16777215,16777215,0,0,0,0,16777215,16777215,16777215,16777215,16777215,16777215,16777215,0,0,0,0,16777215,16777215,16777215,16777215,16777215,16777215,16777215,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1"
Private Hold_BarSpace As Integer

Property Let BarSpace(StrValue As Integer)
    Hold_BarSpace = StrValue
    LoadGUI
End Property

Property Get BarSpace() As Integer
    BarSpace = Hold_BarSpace
End Property

Function LoadShell()
    ImageList1.ListImages.Add 1, , Picture1.Image
    UserControl.Picture = ImageList1.ListImages(1).Picture
    UserControl.MaskPicture = UserControl.Picture
    ImageList1.ListImages.Remove 1
End Function

Private Function LoadBmpMenuLines(Legnth As Integer, ColorPallet As String, X As Integer, Y As Integer, Optional Gray As Boolean = True, Optional Brightness As Integer, Optional IgnoreColor As OLE_COLOR = vbWhite) As Integer
    If ColorPallet = "" Then Exit Function
    Dim PixCount
    Dim Colors() As String, CurrentRow, CurrentColumn, Count, Rows
    Colors = Split(ColorPallet, ",")
    Rows = Int(Split(ColorPallet, ",")(0))
    For Count = 1 To UBound(Colors)
        If CurrentRow > (Rows) Then CurrentRow = 0: CurrentColumn = CurrentColumn + 1
            If Colors(Count) <> -1 And Colors(Count) <> IgnoreColor Then
                    Picture1.Line (X + CurrentColumn, Y + CurrentRow)-(X + CurrentColumn + Legnth, Y + CurrentRow), MakeGrey(Colors(Count))
            End If
        CurrentRow = CurrentRow + 1
    Next
    LoadBmpMenuLines = CurrentColumn
End Function

Function LoadSquares(BarCount As Integer)
    Dim X As Integer, Y
    X = 3
    For Y = 1 To BarCount
        X = X + LoadBmpMenuLines(1, Square_Reg, X, 1) + Hold_BarSpace
    Next
End Function

Function LoadGUI()
Picture1.Cls
LoadSquares 15
LoadShell
End Function


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Hold_BarSpace = PropBag.ReadProperty("Hold_BarSpace", 2)
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

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Hold_BarSpace", Hold_BarSpace, 2
End Sub
