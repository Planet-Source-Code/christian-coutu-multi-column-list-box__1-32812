VERSION 5.00
Begin VB.UserControl CoolListBox 
   BackStyle       =   0  'Transparent
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5115
   PropertyPages   =   "ListBox.ctx":0000
   ScaleHeight     =   6180
   ScaleWidth      =   5115
   ToolboxBitmap   =   "ListBox.ctx":005C
   Begin VB.PictureBox PicBorder 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   5910
      Left            =   180
      ScaleHeight     =   5910
      ScaleWidth      =   4785
      TabIndex        =   0
      Top             =   135
      Width           =   4785
      Begin VB.PictureBox PicContainer 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   5055
         Left            =   315
         ScaleHeight     =   5055
         ScaleWidth      =   4335
         TabIndex        =   1
         Top             =   405
         Width           =   4335
         Begin VB.HScrollBar HScroll1 
            Height          =   240
            LargeChange     =   10
            Left            =   540
            Max             =   0
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   3735
            Visible         =   0   'False
            Width           =   2490
         End
         Begin VB.PictureBox CapPic 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            FillColor       =   &H8000000F&
            Height          =   870
            Left            =   180
            ScaleHeight     =   870
            ScaleWidth      =   4110
            TabIndex        =   6
            Top             =   405
            Visible         =   0   'False
            Width           =   4110
            Begin VB.PictureBox CapSep 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   19
               Left            =   3690
               ScaleHeight     =   285
               ScaleWidth      =   60
               TabIndex        =   26
               Top             =   405
               Visible         =   0   'False
               Width           =   60
            End
            Begin VB.PictureBox CapSep 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   18
               Left            =   3510
               ScaleHeight     =   285
               ScaleWidth      =   60
               TabIndex        =   25
               Top             =   405
               Visible         =   0   'False
               Width           =   60
            End
            Begin VB.PictureBox CapSep 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   17
               Left            =   3330
               ScaleHeight     =   285
               ScaleWidth      =   60
               TabIndex        =   24
               Top             =   405
               Visible         =   0   'False
               Width           =   60
            End
            Begin VB.PictureBox CapSep 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   16
               Left            =   3150
               ScaleHeight     =   285
               ScaleWidth      =   60
               TabIndex        =   23
               Top             =   405
               Visible         =   0   'False
               Width           =   60
            End
            Begin VB.PictureBox CapSep 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   15
               Left            =   2970
               ScaleHeight     =   285
               ScaleWidth      =   60
               TabIndex        =   22
               Top             =   405
               Visible         =   0   'False
               Width           =   60
            End
            Begin VB.PictureBox CapSep 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   14
               Left            =   2790
               ScaleHeight     =   285
               ScaleWidth      =   60
               TabIndex        =   21
               Top             =   405
               Visible         =   0   'False
               Width           =   60
            End
            Begin VB.PictureBox CapSep 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   13
               Left            =   2070
               ScaleHeight     =   285
               ScaleWidth      =   60
               TabIndex        =   20
               Top             =   405
               Visible         =   0   'False
               Width           =   60
            End
            Begin VB.PictureBox CapSep 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   12
               Left            =   2250
               ScaleHeight     =   285
               ScaleWidth      =   60
               TabIndex        =   19
               Top             =   405
               Visible         =   0   'False
               Width           =   60
            End
            Begin VB.PictureBox CapSep 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   11
               Left            =   2430
               ScaleHeight     =   285
               ScaleWidth      =   60
               TabIndex        =   18
               Top             =   405
               Visible         =   0   'False
               Width           =   60
            End
            Begin VB.PictureBox CapSep 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   10
               Left            =   2610
               ScaleHeight     =   285
               ScaleWidth      =   60
               TabIndex        =   17
               Top             =   405
               Visible         =   0   'False
               Width           =   60
            End
            Begin VB.PictureBox CapSep 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   9
               Left            =   1890
               ScaleHeight     =   285
               ScaleWidth      =   60
               TabIndex        =   16
               Top             =   405
               Visible         =   0   'False
               Width           =   60
            End
            Begin VB.PictureBox CapSep 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   8
               Left            =   1710
               ScaleHeight     =   285
               ScaleWidth      =   60
               TabIndex        =   15
               Top             =   405
               Visible         =   0   'False
               Width           =   60
            End
            Begin VB.PictureBox CapSep 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   7
               Left            =   1530
               ScaleHeight     =   285
               ScaleWidth      =   60
               TabIndex        =   14
               Top             =   405
               Visible         =   0   'False
               Width           =   60
            End
            Begin VB.PictureBox CapSep 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   6
               Left            =   1350
               ScaleHeight     =   285
               ScaleWidth      =   60
               TabIndex        =   13
               Top             =   405
               Visible         =   0   'False
               Width           =   60
            End
            Begin VB.PictureBox CapSep 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   5
               Left            =   1170
               ScaleHeight     =   285
               ScaleWidth      =   60
               TabIndex        =   12
               Top             =   405
               Visible         =   0   'False
               Width           =   60
            End
            Begin VB.PictureBox CapSep 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   4
               Left            =   990
               ScaleHeight     =   285
               ScaleWidth      =   60
               TabIndex        =   11
               Top             =   405
               Visible         =   0   'False
               Width           =   60
            End
            Begin VB.PictureBox CapSep 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   3
               Left            =   810
               ScaleHeight     =   285
               ScaleWidth      =   60
               TabIndex        =   10
               Top             =   405
               Visible         =   0   'False
               Width           =   60
            End
            Begin VB.PictureBox CapSep 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   2
               Left            =   630
               ScaleHeight     =   285
               ScaleWidth      =   60
               TabIndex        =   9
               Top             =   405
               Visible         =   0   'False
               Width           =   60
            End
            Begin VB.PictureBox CapSep 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   1
               Left            =   450
               ScaleHeight     =   285
               ScaleWidth      =   60
               TabIndex        =   8
               Top             =   405
               Visible         =   0   'False
               Width           =   60
            End
            Begin VB.PictureBox CapSep 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   0
               Left            =   270
               ScaleHeight     =   285
               ScaleWidth      =   60
               TabIndex        =   7
               Top             =   405
               Visible         =   0   'False
               Width           =   60
            End
         End
         Begin VB.PictureBox HeadPic 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            FillColor       =   &H8000000F&
            Height          =   285
            Left            =   0
            ScaleHeight     =   285
            ScaleWidth      =   1725
            TabIndex        =   5
            Top             =   0
            Width           =   1725
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   2715
            LargeChange     =   10
            Left            =   4095
            Max             =   0
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   990
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox ListPic 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            DrawStyle       =   5  'Transparent
            FillStyle       =   0  'Solid
            Height          =   2670
            Left            =   495
            ScaleHeight     =   2670
            ScaleWidth      =   2265
            TabIndex        =   3
            Top             =   765
            Width           =   2265
            Begin VB.TextBox TxtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   150
               Left            =   180
               TabIndex        =   28
               Top             =   450
               Visible         =   0   'False
               Width           =   1545
            End
         End
         Begin VB.PictureBox BackPic 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1140
            Left            =   2610
            ScaleHeight     =   1140
            ScaleWidth      =   1725
            TabIndex        =   2
            Top             =   45
            Visible         =   0   'False
            Width           =   1725
         End
      End
   End
End
Attribute VB_Name = "CoolListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private Enum mBorderStyle
    mInsert
    mRaised
End Enum

Public Enum mScrollAuto
    Horizontal
    Vertical
    Both
End Enum

Public Enum mBackPicStyle
    Normal
    Tile
    Stretch
End Enum

Public Enum mListBorder
    None
    SmoothInsert
    SmoothRaised
    Insert
    Raised
    Frame
    Groove
End Enum

Dim MyItem() As String
Dim ItemSelected() As Boolean
Dim ItemQty As Long
Dim MaxVQty As Integer
Dim MyMultiSelect As Boolean
Dim MyItemSel As Long
Dim I As Integer
Dim MyListIndex As Long
Dim FocusGot As Boolean
Dim MyScrollAuto As mScrollAuto
Dim MyColSeparator As String
Dim MyHeadCaption As String
Dim MyHeadAlign As AlignmentConstants
Dim MyColWidth As String
Dim MyColCaption As String
Dim MyColCapAlign As String
Dim MyColListAlign As String
Dim MyColCount As Integer
Dim TotalColWidth As Long
Dim MyColIndex As Integer
Dim MyListEdit As Boolean
Dim MyIntegralHeight As Boolean
Dim MyBackPicStyle As mBackPicStyle

Dim MyHeaderBorder As mListBorder
Dim MyColumnBorder As mListBorder
Dim MyListBorder As mListBorder
Dim MyControlBorder As mListBorder

Dim MyTextColor As OLE_COLOR
Dim MySelectedColor As OLE_COLOR
Dim MySelectedTextColor As OLE_COLOR
Dim MyMultiSelTextColor As OLE_COLOR
Dim MyMultiNotSelTextColor As OLE_COLOR

Const defCaption = "Col 1¦Col 2¦Col 3¦Col 4¦Col 5¦Col 6¦Col 7¦Col 8¦Col 9¦Col 10¦Col 11¦Col 12¦Col 13¦Col 14¦Col 15¦Col 16¦Col 17¦Col 18¦Col 19¦Col 20"
Const defWidth = "1000¦1000¦1000¦1000¦1000¦1000¦1000¦1000¦1000¦1000¦1000¦1000¦1000¦1000¦1000¦1000¦1000¦1000¦1000¦1000"
Const defCapAlign = "0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0"
Const defListAlign = "0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0"

Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Property Get HeaderFont() As StdFont
Set HeaderFont = HeadPic.Font
End Property

Public Property Get ColumnFont() As StdFont
Set ColumnFont = CapPic.Font
End Property

Public Property Set ColumnFont(ByVal NewFont As StdFont)
Set CapPic.Font = NewFont
UserControl_Resize
PropertyChanged "ColumnFont"
End Property
Public Property Set HeaderFont(ByVal NewFont As StdFont)
Set HeadPic.Font = NewFont
UserControl_Resize
PropertyChanged "HeaderFont"
End Property
Private Sub Make3D(Obj As Object, BorderWidth As Double, Style As mBorderStyle, Optional LightColor As OLE_COLOR = &H80000016, Optional DarkColor As OLE_COLOR = &H80000010)
Dim DS, DW, SM
Dim K
Dim x2 As Long
Dim y2 As Long

Dim pixX As Integer
Dim pixY As Integer

pixX = Screen.TwipsPerPixelX
pixY = Screen.TwipsPerPixelY

x2 = Obj.Width - pixX
y2 = Obj.Height - pixY

DS = Obj.DrawStyle
DW = Obj.DrawWidth
SM = Obj.ScaleMode

Obj.DrawStyle = 0
Obj.DrawWidth = 1
Obj.ScaleMode = vbTwips

BorderWidth = BorderWidth * 20

If BorderWidth = 0 Then
   BorderWidth = 50
End If

Select Case Style
  Case 0    'Inset
    'Upper border
    For K = 0 To BorderWidth
      Obj.Line (0, 0 + K)-(x2 - K, 0 + K), DarkColor
    Next
    
    'Right border
    For K = 0 To BorderWidth
       Obj.Line (x2 - K, 0 + K)-(x2 - K, y2), LightColor
    Next
    
    'Left border
    For K = 0 To BorderWidth
         Obj.Line (0 + K, 0)-(0 + K, y2 - K), DarkColor
    Next
    
    'Bottom border
    For K = 0 To BorderWidth
      Obj.Line (0 + K, y2 - K)-(x2, y2 - K), LightColor
    Next
  Case 1    'Raised
    'Upper border
    For K = 0 To BorderWidth
      Obj.Line (0, 0 + K)-(x2 - K, 0 + K), LightColor
    Next
    
    'Right border
    For K = 0 To BorderWidth
       Obj.Line (x2 - K, 0 + K)-(x2 - K, y2), DarkColor
    Next
    
    'Left border
    For K = 0 To BorderWidth
         Obj.Line (0 + K, 0)-(0 + K, y2 - K), LightColor
    Next
    
    'Bottom border
    For K = 0 To BorderWidth
      Obj.Line (0 + K, y2 - K)-(x2, y2 - K), DarkColor
    Next

End Select

Obj.DrawStyle = DS          'restore the settings
Obj.DrawWidth = DW
Obj.ScaleMode = SM

End Sub

Private Sub HScroll1_Change()
ListPic.Left = -HScroll1.Value * 100
CapPic.Left = ListPic.Left
DrawColCap
End Sub

Private Sub ListPic_Click()
RaiseEvent Click
End Sub

Private Sub ListPic_DblClick()
RaiseEvent DblClick
End Sub


Private Sub ListPic_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If ListPic.Enabled = True Then
Select Case KeyCode
Case 13
    TxtEdit.Visible = False
    SaveTxtEdit
Case 27
    TxtEdit.Visible = False
Case 32
    If MyMultiSelect = True Then
    ItemSelected(MyListIndex) = Not ItemSelected(MyListIndex)
    RefreshList
    End If
Case 37
    If HScroll1.Value <> 0 Then HScroll1.Value = HScroll1.Value - 1
Case 38
    Me.ListIndex = Me.ListIndex - 1
Case 39
    If HScroll1.Value <> HScroll1.Max Then HScroll1.Value = HScroll1.Value + 1
Case 40
    If MyListIndex < ItemQty - 1 Then Me.ListIndex = Me.ListIndex + 1
End Select
RaiseEvent KeyDown(KeyCode, Shift)
End If
End Sub

Private Sub ListPic_KeyPress(KeyAscii As Integer)
If ListPic.Enabled = True Then RaiseEvent KeyPress(KeyAscii)
End Sub


Private Sub ListPic_KeyUp(KeyCode As Integer, Shift As Integer)
If ListPic.Enabled = True Then RaiseEvent KeyUp(KeyCode, Shift)
End Sub


Private Sub ListPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim yPos As Long
Dim mPos As Integer
Dim TxtH As Long

TxtH = ListPic.TextHeight("X")
mPos = (Y - (TxtH / 2)) / TxtH
yPos = mPos * TxtH

If Button = 1 And TxtEdit.Visible = False And ListPic.Enabled = True Then
MyColIndex = FindCol(X)
If mPos < ItemQty Then
MyListIndex = mPos
    If MyMultiSelect = True Then
    ItemSelected(mPos) = Not ItemSelected(mPos)
    End If
RefreshList
End If
ElseIf Button = 2 And MyListIndex > -1 And MyListEdit = True Then
MyColIndex = FindCol(X)
TxtEdit.BackColor = ListPic.BackColor
TxtEdit.ForeColor = MyTextColor
TxtEdit.Visible = True
TxtEdit.Height = ListPic.TextHeight("X") + 20
TxtEdit.Top = MyListIndex * ListPic.TextHeight("X")
    If MyColIndex = MyColCount - 1 Then
        If VScroll1.Visible = False Then
        TxtEdit.Width = CapPic.Width - CapSep(MyColIndex - 1).Left
        Else
        TxtEdit.Width = CapPic.Width - CapSep(MyColIndex - 1).Left - VScroll1.Width
        End If
    Else
    TxtEdit.Width = Split(MyColWidth, "¦")(MyColIndex)
    End If
    If MyColIndex = 0 Then
    TxtEdit.Left = 0
    Else
    TxtEdit.Left = CapSep(MyColIndex - 1).Left + 30
    End If
TxtEdit.Text = Split(MyItem(MyListIndex), MyColSeparator)(MyColIndex)
TxtEdit.Alignment = Split(MyColListAlign, "¦")(MyColIndex)
TxtEdit.SelStart = 0
TxtEdit.SelLength = Len(TxtEdit.Text)
TxtEdit.SetFocus
End If
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub


Private Sub ListPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ListPic.Enabled = True Then RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub


Private Sub ListPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ListPic.Enabled = True Then RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub


Private Sub TxtEdit_GotFocus()
TxtEdit.SelStart = 0
TxtEdit.SelLength = Len(TxtEdit.Text)
TxtEdit.SetFocus
End Sub

Private Sub TxtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
TxtEdit.Visible = False
SaveTxtEdit
ElseIf KeyCode = 27 Then
TxtEdit.Visible = False
End If
End Sub

Private Sub UserControl_Initialize()
MyListIndex = -1
End Sub

Private Sub UserControl_InitProperties()
MyMultiSelect = False
MyScrollAuto = 2
MyColSeparator = "|"
MyHeadCaption = Ambient.DisplayName
MyColCaption = defCaption
MyColWidth = defWidth
MyColCapAlign = defCapAlign
MyColListAlign = defListAlign
MyHeadAlign = vbCenter
MyColCount = 1
ListPic.Enabled = True

MyTextColor = &H80000008
ListPic.BackColor = &H80000005
MySelectedColor = &H8000000D
SelectedTextColor = &H8000000E
HeaderBackColor = &H8000000F
MultiSelTextColor = vbYellow
MultiNotSelTextColor = vbBlue
MyListEdit = False

MyColumnBorder = SmoothRaised
MyHeaderBorder = SmoothRaised
MyListBorder = SmoothInsert
MyControlBorder = SmoothInsert

DrawHeadCaption
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim Indx As Integer
With PropBag
ListPic.Enabled = .ReadProperty("Enabled", True)
MyIntegralHeight = .ReadProperty("IntegralHeight", True)
MyMultiSelect = .ReadProperty("MultiSelect", True)
MyListEdit = .ReadProperty("EditableList", False)
MyScrollAuto = .ReadProperty("ScrollBarAuto", 1)
MyColSeparator = .ReadProperty("ColSeparator", "|")
MyHeadCaption = .ReadProperty("HeaderCaption", Ambient.DisplayName)
MyHeadAlign = .ReadProperty("HeaderAlignment", 2)
MyColCount = .ReadProperty("ColumnCount", 1)
MyColWidth = .ReadProperty("ColumnWidth", 0)
MyColCaption = .ReadProperty("ColumnCaption", "")
MyColCapAlign = .ReadProperty("ColCaptionAlignment", "")
MyColListAlign = .ReadProperty("ColListAlignment", "")

Set BackPic.Picture = .ReadProperty("BackPicture", Nothing)
MyBackPicStyle = .ReadProperty("BackPicStyle", 0)

ListPic.BackColor = .ReadProperty("BackColor", &H80000005)
MyTextColor = .ReadProperty("TextColor", &H80000008)
MySelectedColor = .ReadProperty("SelectedColor", &H8000000D)
MySelectedTextColor = .ReadProperty("SelectedTextColor", &H8000000E)
HeadPic.BackColor = .ReadProperty("HeaderBackColor", &H8000000F)
HeadPic.ForeColor = .ReadProperty("HeaderForeColor", &H80000012)
CapPic.BackColor = .ReadProperty("ColumnBackColor", &H8000000F)
CapPic.ForeColor = .ReadProperty("ColumnForeColor", &H80000012)
MyMultiSelTextColor = .ReadProperty("MultiSelTextColor", vbYellow)
MyMultiNotSelTextColor = .ReadProperty("MultiNotSelTextColor", vbBlue)
For I = 0 To MyColCount
CapSep(I).BackColor = .ReadProperty("ColumnBackColor", &H8000000F)
Next I

MyHeaderBorder = .ReadProperty("HeaderBorder", 2)
MyColumnBorder = .ReadProperty("ColumnBorder", 2)
MyListBorder = .ReadProperty("ListBorder", 1)
MyControlBorder = .ReadProperty("ControlBorder", 1)

Set ListPic.Font = .ReadProperty("ListFont", Ambient.Font)
Set HeadPic.Font = .ReadProperty("HeaderFont", Ambient.Font)
Set CapPic.Font = .ReadProperty("ColumnFont", Ambient.Font)
Set TxtEdit.Font = .ReadProperty("ListFont", Ambient.Font)
End With

End Sub


Private Sub UserControl_Resize()
On Error Resume Next
Dim TmpHeight As Long
Dim TmpColW As Long
Dim TotW As Long
Dim Th As Long
Dim K As Integer

CapSep(0).Left = PicContainer.Width

PicBorder.Left = 0
PicBorder.Top = 0
PicBorder.Height = UserControl.Height
PicBorder.Width = UserControl.Width

PicContainer.Left = 30
PicContainer.Top = 30
PicContainer.Height = UserControl.Height - 60
PicContainer.Width = UserControl.Width - 60

If Len(MyHeadCaption) > 0 Then
HeadPic.Height = HeadPic.TextHeight("X") + 50
Else
HeadPic.Height = 0
End If
CapPic.Visible = True
CapPic.Height = CapPic.TextHeight("X") + 50

MaxVQty = (PicContainer.Height - HeadPic.Height - CapPic.Height) / ListPic.TextHeight("X")

HeadPic.Top = 0
HeadPic.Left = 0
HeadPic.Width = PicContainer.Width ' - 10
CapPic.Top = HeadPic.Height
CapPic.Left = 0
CapPic.Width = PicContainer.Width

For K = 0 To 19
CapSep(K).Visible = False
Next K

For K = 0 To MyColCount - 2
CapSep(K).Top = 0
CapSep(K).Height = CapPic.Height
TmpColW = CLng(Split(MyColWidth, "¦")(K))
TotW = TotW + TmpColW
CapSep(K).Left = TotW
CapSep(K).Visible = True
Next K
ListPic.Width = PicContainer.Width ' - 10
ListPic.Top = HeadPic.Height + CapPic.Height
ListPic.Left = 0
DrawColCap
ResizeList
If Len(MyHeadCaption) > 0 Then DrawHdCap
End Sub



Public Sub Additem(Item As String, Optional Index As Integer = -1)
Dim J As Integer
Dim K As Integer
Dim TmpX As Long
Dim TotX As Long
Dim TmpY As Long

If Index = -1 Then
ReDim Preserve MyItem(ItemQty + 1)
ReDim Preserve ItemSelected(ItemQty + 1)
MyItem(ItemQty) = Item
ItemQty = ItemQty + 1
Else
ReDim Preserve MyItem(ItemQty + 1)
ReDim Preserve ItemSelected(ItemQty + 1)
For J = ItemQty + 1 To Index + 1 Step -1
MyItem(J) = MyItem(J - 1)
    If MyMultiSelect = True Then
    ItemSelected(J) = ItemSelected(J - 1)
    End If
Next J
MyItem(Index) = Item
ItemSelected(Index) = False
ItemQty = ItemQty + 1
MyListIndex = Index
End If
HScroll1.Value = 0
End Sub

Public Sub ReplaceItem(IndexToReplace As Long, NewItem As String)
If IndexToReplace <> -1 Then
MyItem(IndexToReplace) = NewItem
MyListIndex = IndexToReplace
RefreshList
End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim Indx As Integer
With PropBag
.WriteProperty "Enabled", ListPic.Enabled, True
.WriteProperty "IntegralHeight", MyIntegralHeight, True
.WriteProperty "MultiSelect", MyMultiSelect, True
.WriteProperty "EditableList", MyListEdit, False
.WriteProperty "ScrollBarAuto", MyScrollAuto, 1
.WriteProperty "ColSeparator", MyColSeparator, "|"
.WriteProperty "HeaderCaption", MyHeadCaption, Ambient.DisplayName
.WriteProperty "HeaderAlignment", MyHeadAlign, 2
.WriteProperty "ColumnCount", MyColCount, 1
.WriteProperty "ColumnWidth", MyColWidth, 0
.WriteProperty "ColumnCaption", MyColCaption, ""
.WriteProperty "ColCaptionAlignment", MyColCapAlign, ""
.WriteProperty "ColListAlignment", MyColListAlign, ""

.WriteProperty "BackPicture", BackPic.Picture, Nothing
.WriteProperty "BackPicStyle", MyBackPicStyle, 0

.WriteProperty "BackColor", ListPic.BackColor, &H80000005
.WriteProperty "TextColor", MyTextColor, &H80000008
.WriteProperty "SelectedColor", MySelectedColor, &H8000000D
.WriteProperty "SelectedTextColor", MySelectedTextColor, &H8000000E
.WriteProperty "HeaderBackColor", HeadPic.BackColor, &H8000000F
.WriteProperty "HeaderForeColor", HeadPic.ForeColor, &H80000012
.WriteProperty "ColumnBackColor", CapPic.BackColor, &H8000000F
.WriteProperty "ColumnForeColor", CapPic.ForeColor, &H80000012
.WriteProperty "MultiSelTextColor", MyMultiSelTextColor, vbYellow
.WriteProperty "MultiNotSelTextColor", MyMultiNotSelTextColor, vbBlue
For I = 0 To MyColCount
.WriteProperty "ColumnBackColor", CapSep(I).BackColor, &H8000000F
Next I

.WriteProperty "HeaderBorder", MyHeaderBorder, 2
.WriteProperty "ColumnBorder", MyColumnBorder, 2
.WriteProperty "ListBorder", MyListBorder, 1
.WriteProperty "ControlBorder", MyControlBorder, 1

.WriteProperty "ListFont", ListPic.Font, Ambient.Font
.WriteProperty "HeaderFont", HeadPic.Font, Ambient.Font
.WriteProperty "ColumnFont", CapPic.Font, Ambient.Font
.WriteProperty "ListFont", TxtEdit.Font, Ambient.Font
End With

End Sub

Private Sub VScroll1_Change()
ListPic.Top = (-(VScroll1.Value * ListPic.TextHeight("X"))) + HeadPic.Height + CapPic.Height
End Sub



Public Sub Clear()
ReDim MyItem(0)
ReDim ItemSelected(0)
ItemQty = 0
ListPic.CurrentX = 0
ListPic.CurrentY = 0
ListPic.Cls
If HScroll1.Visible = False Then
ListPic.Height = PicContainer.Height - CapPic.Height - HeadPic.Height
Else
ListPic.Height = PicContainer.Height - CapPic.Height - HeadPic.Height - HScroll1.Height
End If
ResizeList
MyListIndex = -1
HScroll1.Value = 0
VScroll1.Value = 0
TxtEdit.Text = "2"
TxtEdit.Visible = False
DrawColCap
End Sub


Public Property Get Selected(Index As Integer) As Boolean
Attribute Selected.VB_MemberFlags = "400"
Selected = ItemSelected(Index)
End Property

Public Property Let Selected(Index As Integer, ByVal vNewValue As Boolean)
ItemSelected(Index) = vNewValue
End Property


Public Property Get MultiSelect() As Boolean
Attribute MultiSelect.VB_ProcData.VB_Invoke_Property = "General"
MultiSelect = MyMultiSelect
End Property

Public Property Let MultiSelect(ByVal vNewValue As Boolean)
MyMultiSelect = vNewValue
RefreshList
End Property

Public Sub RemoveItem(Index As Integer)
Dim J
TxtEdit.Visible = False
If Index >= 0 And Index < ItemQty Then
    If Index < ItemQty - 1 Then
    For J = Index To ItemQty - 1
    MyItem(J) = MyItem(J + 1)
    Next J
    ItemQty = ItemQty - 1
    ReDim Preserve MyItem(ItemQty)
    ReDim Preserve ItemSelected(ItemQty)
    'RefreshList
    ElseIf Index = ItemQty - 1 Then
    ItemQty = ItemQty - 1
    ReDim Preserve MyItem(ItemQty)
    ReDim Preserve ItemSelected(ItemQty)
    'RefreshList
    MyListIndex = -1
    End If
End If
End Sub

Private Sub ResizeList()

CountMaxWidth

Select Case MyScrollAuto
Case 0
    If TotalColWidth < PicContainer.Width Then
    HScroll1.Visible = False
    ListPic.Width = PicContainer.Width
    CapPic.Width = PicContainer.Width
        If MyIntegralHeight = False Then
        ListPic.Height = PicContainer.Height - HeadPic.Height - CapPic.Height
        Else
        ListPic.Height = MaxVQty * ListPic.TextHeight("X")
        UserControl.Height = HeadPic.Height + CapPic.Height + ListPic.Height + 60
        End If
    Else
    HScroll1.Visible = True
    ListPic.Width = TotalColWidth - 20
    CapPic.Width = TotalColWidth - 50
        If MyIntegralHeight = False Then
        ListPic.Height = PicContainer.Height - HeadPic.Height - CapPic.Height - HScroll1.Height
        Else
        ListPic.Height = (MaxVQty - 1) * ListPic.TextHeight("X")
        UserControl.Height = HeadPic.Height + CapPic.Height + ListPic.Height + HScroll1.Height + 60
        End If
    HScroll1.Left = 0
    HScroll1.Top = PicContainer.Height - HScroll1.Height
    HScroll1.Width = PicContainer.Width
    HScroll1.Max = (TotalColWidth - PicContainer.Width) / 100
    End If
Case 1
    If ItemQty > MaxVQty Then
    VScroll1.Visible = True
    ListPic.Width = PicContainer.Width - VScroll1.Width
    ListPic.Height = ItemQty * ListPic.TextHeight("X")
    VScroll1.Left = PicContainer.Width - VScroll1.Width
    VScroll1.Top = HeadPic.Height + CapPic.Height
    VScroll1.Height = PicContainer.Height - HeadPic.Height - CapPic.Height
    VScroll1.Max = (ListPic.Height - PicContainer.Height + CapPic.Height + HeadPic.Height + 60) / ListPic.TextHeight("X")
    Else
    VScroll1.Visible = False
    ListPic.Width = PicContainer.Width
    ListPic.Top = HeadPic.Height + CapPic.Height
        If MyIntegralHeight = False Then
            If ListPic.Height < (PicContainer.Height - CapPic.Height - HeadPic.Height) Then
            ListPic.Height = PicContainer.Height - CapPic.Height - HeadPic.Height
            Else
            ListPic.Height = PicContainer.Height - HeadPic.Height - CapPic.Height
            End If
        Else
        ListPic.Height = MaxVQty * ListPic.TextHeight("X")
        UserControl.Height = HeadPic.Height + CapPic.Height + ListPic.Height + 60
        End If
    End If
Case 2
    If TotalColWidth > PicContainer.Width Then
    HScroll1.Visible = True
    HScroll1.Left = 0
    HScroll1.Top = PicContainer.Height - HScroll1.Height
    HScroll1.Width = PicContainer.Width
    HScroll1.Max = (TotalColWidth - PicContainer.Width) / 100
    CapPic.Width = TotalColWidth - 50
        If ItemQty > MaxVQty Then
        VScroll1.Visible = True
        ListPic.Width = TotalColWidth - VScroll1.Width - 20
        ListPic.Height = ItemQty * ListPic.TextHeight("X")
        VScroll1.Left = PicContainer.Width - VScroll1.Width
        VScroll1.Top = HeadPic.Height + CapPic.Height
        VScroll1.Height = PicContainer.Height - HeadPic.Height - CapPic.Height - HScroll1.Height
        VScroll1.Max = (ListPic.Height - PicContainer.Height + CapPic.Height + HeadPic.Height + HScroll1.Height) / ListPic.TextHeight("X") ' + 60
        Else
        VScroll1.Visible = False
        ListPic.Width = TotalColWidth - 20
        ListPic.Top = HeadPic.Height + CapPic.Height
        If MyIntegralHeight = False Then
                If ListPic.Height < (PicContainer.Height - CapPic.Height - HeadPic.Height) Then
                ListPic.Height = PicContainer.Height - CapPic.Height - HeadPic.Height
                Else
                ListPic.Height = PicContainer.Height - HeadPic.Height - CapPic.Height
                End If
            Else
            ListPic.Height = (MaxVQty - 1) * ListPic.TextHeight("X")
            UserControl.Height = HeadPic.Height + CapPic.Height + ListPic.Height + HScroll1.Height + 60
            End If
        End If
    Else
    HScroll1.Visible = False
    HScroll1.Top = PicContainer.Height
    ListPic.Width = PicContainer.Width
    CapPic.Width = PicContainer.Width
        If ItemQty > MaxVQty Then
        VScroll1.Visible = True
        ListPic.Height = ItemQty * ListPic.TextHeight("X")
            If ListPic.Height < (PicContainer.Height - CapPic.Height - HeadPic.Height) Then
            ListPic.Height = PicContainer.Height - CapPic.Height - HeadPic.Height
            End If
        VScroll1.Left = PicContainer.Width - VScroll1.Width
        VScroll1.Top = HeadPic.Height + CapPic.Height
        VScroll1.Height = PicContainer.Height - HeadPic.Height - CapPic.Height
        VScroll1.Max = (ListPic.Height - PicContainer.Height + CapPic.Height + HeadPic.Height) / ListPic.TextHeight("X") ' + 120
        Else
        VScroll1.Visible = False
        ListPic.Top = HeadPic.Height + CapPic.Height
            If MyIntegralHeight = False Then
                If ListPic.Height < (PicContainer.Height - CapPic.Height - HeadPic.Height) Then
                ListPic.Height = PicContainer.Height - CapPic.Height - HeadPic.Height
                Else
                ListPic.Height = PicContainer.Height - HeadPic.Height - CapPic.Height
                End If
            Else
            ListPic.Height = MaxVQty * ListPic.TextHeight("X")
            UserControl.Height = HeadPic.Height + CapPic.Height + ListPic.Height + 60
            End If
        End If
        If VScroll1.Visible = True Then
        ListPic.Width = PicContainer.Width - VScroll1.Width
        Else
        ListPic.Width = PicContainer.Width
        End If
    End If
    If ListPic.Height < (PicContainer.Height - CapPic.Height - HeadPic.Height) Then
    ListPic.Height = PicContainer.Height - CapPic.Height - HeadPic.Height
    End If
End Select
RefreshList
End Sub

Public Property Get ListCount() As Long
ListCount = ItemQty
End Property


Public Property Get ListIndex() As Long
Attribute ListIndex.VB_MemberFlags = "400"
ListIndex = MyListIndex
End Property

Public Property Let ListIndex(ByVal vNewValue As Long)
If vNewValue > -1 And vNewValue < ItemQty Then
MyListIndex = vNewValue
End If
If HScroll1.Visible = True Then
    If MyListIndex > MaxVQty - 3 And MyListIndex < ItemQty Then
    VScroll1.Value = MyListIndex - MaxVQty + 2
    End If
Else
    If MyListIndex > MaxVQty - 2 And MyListIndex < ItemQty Then
    VScroll1.Value = MyListIndex - MaxVQty + 1
    End If
End If
RefreshList
End Property

Public Property Get ScrollBarAuto() As mScrollAuto
ScrollBarAuto = MyScrollAuto
End Property

Public Property Let ScrollBarAuto(ByVal vNewValue As mScrollAuto)
MyScrollAuto = vNewValue
PropertyChanged "ScrollBarAuto"
End Property

Public Property Get ColSeparator() As String
Attribute ColSeparator.VB_ProcData.VB_Invoke_Property = "General"
ColSeparator = MyColSeparator
End Property

Public Property Let ColSeparator(ByVal vNewValue As String)
MyColSeparator = vNewValue
PropertyChanged "ColSeparator"
End Property

Public Property Get HeaderCaption() As String
Attribute HeaderCaption.VB_ProcData.VB_Invoke_Property = "General"
HeaderCaption = MyHeadCaption
End Property

Public Property Let HeaderCaption(ByVal vNewValue As String)
MyHeadCaption = vNewValue
DrawHeadCaption
PropertyChanged "HeaderCaption"
End Property

Public Property Get HeaderAlignment() As AlignmentConstants
HeaderAlignment = MyHeadAlign
End Property

Public Property Let HeaderAlignment(ByVal vNewValue As AlignmentConstants)
MyHeadAlign = vNewValue
DrawHeadCaption
PropertyChanged "HeaderAlignment"
End Property

Private Sub DrawHeadCaption()
DrawHdCap
UserControl_Resize

End Sub

Private Sub DrawHdCap()
HeadPic.CurrentX = 0
HeadPic.CurrentY = 0
HeadPic.Cls

DrawBorder
If Len(MyHeadCaption) > 0 Then
Select Case MyHeadAlign
Case 0
HeadPic.CurrentX = 50
HeadPic.CurrentY = (HeadPic.Height - HeadPic.TextHeight("X")) / 2
HeadPic.Print MyHeadCaption
Case 1
HeadPic.CurrentX = HeadPic.Width - HeadPic.TextWidth(MyHeadCaption) - 50
HeadPic.CurrentY = (HeadPic.Height - HeadPic.TextHeight("X")) / 2
HeadPic.Print MyHeadCaption
Case 2
HeadPic.CurrentX = (HeadPic.Width - HeadPic.TextWidth(MyHeadCaption)) / 2
HeadPic.CurrentY = (HeadPic.Height - HeadPic.TextHeight("X")) / 2
HeadPic.Print MyHeadCaption
End Select
End If
End Sub

Private Sub DrawColCap()
Dim TmpCap As String
Dim TmpAlign As Integer

CapPic.CurrentX = 0
CapPic.CurrentY = 0
CapPic.Cls

DrawBorder
For I = 0 To MyColCount - 1
TmpAlign = Int(Split(MyColCapAlign, "¦")(I))
TmpCap = Split(MyColCaption, "¦")(I)
Select Case TmpAlign
Case 0
    If I = 0 Then
    CapPic.CurrentX = 30
    Else
    CapPic.CurrentX = CapSep(I - 1).Left + 60
    End If
Case 1
    If I = 0 Then
    CapPic.CurrentX = (CapSep(I).Left - CapPic.TextWidth(TmpCap)) - 60
    ElseIf I = MyColCount - 1 Then
    CapPic.CurrentX = CapPic.Width - CapPic.TextWidth(TmpCap) - 60
    Else
    CapPic.CurrentX = CapSep(I).Left - CapPic.TextWidth(TmpCap) - 60
    End If
Case 2
    If I = 0 Then
    CapPic.CurrentX = (CapSep(I).Left - CapPic.TextWidth(TmpCap)) / 2
    ElseIf I = MyColCount - 1 Then
    CapPic.CurrentX = (CapPic.Width - CapPic.TextWidth(TmpCap) + CapSep(I - 1).Left) / 2
    Else
    CapPic.CurrentX = (CapSep(I).Left - CapPic.TextWidth(TmpCap) + CapSep(I - 1).Left) / 2
    End If
End Select

CapPic.CurrentY = (CapPic.Height - CapPic.TextHeight("X")) / 2
CapPic.Print TmpCap
Next I

End Sub
Public Property Get ColumnWidth(Index As Integer) As Long
Attribute ColumnWidth.VB_MemberFlags = "400"
ColumnWidth = CLng(Split(MyColWidth, "¦")(Index))
End Property

Public Property Let ColumnWidth(Index As Integer, ByVal vNewValue As Long)
SetColWidth Index, vNewValue
UserControl_Resize
PropertyChanged "ColumnWidth"
End Property

Public Property Get ColumnCaption(Index As Integer) As String
Attribute ColumnCaption.VB_ProcData.VB_Invoke_Property = "Column"
ColumnCaption = Split(MyColCaption, "¦")(Index)
End Property

Public Property Let ColumnCaption(Index As Integer, ByVal vNewValue As String)
SetColCaption Index, vNewValue
UserControl_Resize
PropertyChanged "ColumnCaption"
End Property

Public Property Get ColumnCount() As Integer
Attribute ColumnCount.VB_ProcData.VB_Invoke_Property = "Column"
ColumnCount = MyColCount
End Property

Public Property Let ColumnCount(ByVal vNewValue As Integer)
If vNewValue > 20 Then
MsgBox "This ListBox is limited to 20 columns"
Else
MyColCount = vNewValue
UserControl_Resize
End If

PropertyChanged "ColumnCount"
End Property

Private Sub SetColCaption(Index As Integer, mCaption As String)
On Error Resume Next
Dim OldCap(19) As String
For I = 0 To 19
OldCap(I) = Split(MyColCaption, "¦")(I)
Next I
OldCap(Index) = mCaption
MyColCaption = ""
For I = 0 To 19
MyColCaption = MyColCaption & OldCap(I) & "¦"
Next I
End Sub
Private Sub SetColWidth(Index As Integer, mColWidth As Long)
On Error Resume Next
Dim OldWidth(19) As String

For I = 0 To 19
OldWidth(I) = Split(MyColWidth, "¦")(I)
Next I
OldWidth(Index) = mColWidth
MyColWidth = ""

For I = 0 To 19
MyColWidth = MyColWidth & OldWidth(I) & "¦"
Next I
End Sub



Public Property Get ColCaptionAlignment(Index As Integer) As AlignmentConstants
ColCaptionAlignment = Split(MyColCapAlign, "¦")(Index)
End Property

Public Property Get ColListAlignment(Index As Integer) As AlignmentConstants
ColListAlignment = Split(MyColListAlign, "¦")(Index)
End Property
Public Property Let ColCaptionAlignment(Index As Integer, ByVal vNewValue As AlignmentConstants)
SetColCapAlign Index, vNewValue
UserControl_Resize
PropertyChanged "ColCaptionAlignment"
End Property

Public Property Let ColListAlignment(Index As Integer, ByVal vNewValue As AlignmentConstants)
SetColListAlign Index, vNewValue
ResizeList
PropertyChanged "ColListAlignment"
End Property
Private Sub SetColCapAlign(Index As Integer, mCapAlign As AlignmentConstants)
On Error Resume Next
Dim OldAlign(19) As String
For I = 0 To 19
OldAlign(I) = Split(MyColCapAlign, "¦")(I)
Next I
OldAlign(Index) = Trim(Str(mCapAlign))
MyColCapAlign = ""
For I = 0 To 19
MyColCapAlign = MyColCapAlign & OldAlign(I) & "¦"
Next I
End Sub
Private Sub SetColListAlign(Index As Integer, mListAlign As AlignmentConstants)
On Error Resume Next
Dim OldAlign(19) As String
For I = 0 To 19
OldAlign(I) = Split(MyColCapAlign, "¦")(I)
Next I
OldAlign(Index) = Trim(Str(mListAlign))
MyColListAlign = ""
For I = 0 To 19
MyColListAlign = MyColListAlign & OldAlign(I) & "¦"
Next I
End Sub

Private Sub CountMaxWidth()
TotalColWidth = 0
For I = 0 To MyColCount - 1
TotalColWidth = TotalColWidth + CLng(Split(MyColWidth, "¦")(I))
Next I
TotalColWidth = TotalColWidth + 50
End Sub

Private Sub RefreshList()
On Error Resume Next
Dim K As Integer
Dim TmpX As Long
Dim TotX As Long
Dim TmpY As Long
Dim TmpAlign As Integer
Dim TmpCap As String

ListPic.CurrentX = 0
ListPic.CurrentY = 0
ListPic.Cls
TmpX = 15
DrawBackPic
If MyMultiSelect = False Then
    For I = 0 To ItemQty - 1
    TmpY = I * ListPic.TextHeight("X")
    If I = MyListIndex Then
        ListPic.FillColor = MySelectedColor
        ListPic.ForeColor = MySelectedTextColor
        ListPic.Line (0, TmpY - 10)-Step(ListPic.ScaleWidth, ListPic.TextHeight("X") + 20), , B
        For K = 0 To MyColCount - 1
        TmpCap = ""
        TmpCap = Split(MyItem(I), MyColSeparator, MyColCount)(K)
        TmpAlign = Int(Split(MyColListAlign, "¦")(K))
        Select Case TmpAlign
        Case 0
            If K = 0 Then
            TotX = 30
            Else
            TotX = CapSep(K - 1).Left + 30
            End If
        Case 1
            If K = 0 Then
            TotX = (CapSep(K).Left - ListPic.TextWidth(TmpCap)) - 30
            ElseIf K = MyColCount - 1 Then
            TotX = ListPic.Width - ListPic.TextWidth(TmpCap) - 50
            Else
            TotX = CapSep(K).Left - ListPic.TextWidth(TmpCap) - 30
            End If
        Case 2
            If K = 0 Then
            TotX = (CapSep(K).Left - ListPic.TextWidth(TmpCap)) / 2
            ElseIf K = MyColCount - 1 Then
            TotX = (ListPic.Width - ListPic.TextWidth(TmpCap) + CapSep(K - 1).Left) / 2
            Else
            TotX = (CapSep(K).Left - ListPic.TextWidth(TmpCap) + CapSep(K - 1).Left) / 2
            End If
        End Select
        ListPic.CurrentX = TotX
        ListPic.CurrentY = TmpY
        ListPic.Print TmpCap
        Next K
    Else
        ListPic.ForeColor = MyTextColor
        For K = 0 To MyColCount - 1
        TmpCap = ""
        TmpCap = Split(MyItem(I), MyColSeparator, MyColCount)(K)
        TmpAlign = Int(Split(MyColListAlign, "¦")(K))
        Select Case TmpAlign
        Case 0
            If K = 0 Then
            TotX = 30
            Else
            TotX = CapSep(K - 1).Left + 30
            End If
        Case 1
            If K = 0 Then
            TotX = (CapSep(K).Left - ListPic.TextWidth(TmpCap)) - 30
            ElseIf K = MyColCount - 1 Then
            TotX = ListPic.Width - ListPic.TextWidth(TmpCap) - 50
            Else
            TotX = CapSep(K).Left - ListPic.TextWidth(TmpCap) - 30
            End If
        Case 2
            If K = 0 Then
            TotX = (CapSep(K).Left - ListPic.TextWidth(TmpCap)) / 2
            ElseIf K = MyColCount - 1 Then
            TotX = (ListPic.Width - ListPic.TextWidth(TmpCap) + CapSep(K - 1).Left) / 2
            Else
            TotX = (CapSep(K).Left - ListPic.TextWidth(TmpCap) + CapSep(K - 1).Left) / 2
            End If
        End Select
        ListPic.CurrentX = TotX
        ListPic.CurrentY = TmpY
        ListPic.Print TmpCap
        Next K
    End If
    Next I
Else

    For I = 0 To ItemQty - 1
    TmpY = I * ListPic.TextHeight("X")
    If ItemSelected(I) = True Then
        ListPic.FillColor = MySelectedColor
            If I = MyListIndex Then
            ListPic.ForeColor = MultiSelTextColor
            Else
            ListPic.ForeColor = MySelectedTextColor
            End If
        ListPic.Line (0, TmpY - 10)-Step(ListPic.ScaleWidth, ListPic.TextHeight("X") + 20), , B
        For K = 0 To MyColCount - 1
        TmpCap = ""
        TmpCap = Split(MyItem(I), MyColSeparator, MyColCount)(K)
        TmpAlign = Int(Split(MyColListAlign, "¦")(K))
        Select Case TmpAlign
        Case 0
            If K = 0 Then
            TotX = 30
            Else
            TotX = CapSep(K - 1).Left + 30
            End If
        Case 1
            If K = 0 Then
            TotX = (CapSep(K).Left - ListPic.TextWidth(TmpCap)) - 30
            ElseIf K = MyColCount - 1 Then
            TotX = ListPic.Width - ListPic.TextWidth(TmpCap) - 50
            Else
            TotX = CapSep(K).Left - ListPic.TextWidth(TmpCap) - 30
            End If
        Case 2
            If K = 0 Then
            TotX = (CapSep(K).Left - ListPic.TextWidth(TmpCap)) / 2
            ElseIf K = MyColCount - 1 Then
            TotX = (ListPic.Width - ListPic.TextWidth(TmpCap) + CapSep(K - 1).Left) / 2
            Else
            TotX = (CapSep(K).Left - ListPic.TextWidth(TmpCap) + CapSep(K - 1).Left) / 2
            End If
        End Select
        ListPic.CurrentX = TotX
        ListPic.CurrentY = TmpY
        ListPic.Print TmpCap
        Next K
    Else
        If I = ListIndex Then
        ListPic.ForeColor = MultiNotSelTextColor
        Else
        ListPic.ForeColor = MyTextColor
        End If
        For K = 0 To MyColCount - 1
        TmpCap = ""
        TmpCap = Split(MyItem(I), MyColSeparator, MyColCount)(K)
        TmpAlign = Int(Split(MyColListAlign, "¦")(K))
        Select Case TmpAlign
        Case 0
            If K = 0 Then
            TotX = 30
            Else
            TotX = CapSep(K - 1).Left + 30
            End If
        Case 1
            If K = 0 Then
            TotX = (CapSep(K).Left - ListPic.TextWidth(TmpCap)) - 30
            ElseIf K = MyColCount - 1 Then
            TotX = ListPic.Width - ListPic.TextWidth(TmpCap) - 50
            Else
            TotX = CapSep(K).Left - ListPic.TextWidth(TmpCap) - 30
            End If
        Case 2
            If K = 0 Then
            TotX = (CapSep(K).Left - ListPic.TextWidth(TmpCap)) / 2
            ElseIf K = MyColCount - 1 Then
            TotX = (ListPic.Width - ListPic.TextWidth(TmpCap) + CapSep(K - 1).Left) / 2
            Else
            TotX = (CapSep(K).Left - ListPic.TextWidth(TmpCap) + CapSep(K - 1).Left) / 2
            End If
        End Select
        ListPic.CurrentX = TotX
        ListPic.CurrentY = TmpY
        ListPic.Print TmpCap
        Next K
    End If
    Next I
End If
DrawColCap
End Sub

Public Property Get BackPicture() As StdPicture
Set BackPicture = BackPic.Picture
End Property

Public Property Set BackPicture(ByVal vNewValue As StdPicture)
Set BackPic.Picture = vNewValue
DrawBackPic
PropertyChanged "BackPicture"
End Property

Public Property Get BackColor() As OLE_COLOR
BackColor = ListPic.BackColor
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
ListPic.BackColor = vNewValue
TxtEdit.BackColor = vNewValue
PicContainer.BackColor = vNewValue
PropertyChanged "BackColor"
End Property

Public Property Get TextColor() As OLE_COLOR
TextColor = MyTextColor
End Property

Public Property Let TextColor(ByVal vNewValue As OLE_COLOR)
MyTextColor = vNewValue
TxtEdit.ForeColor = vNewValue
PropertyChanged "TextColor"
End Property

Public Property Get SelectedColor() As OLE_COLOR
SelectedColor = MySelectedColor
End Property

Public Property Let SelectedColor(ByVal vNewValue As OLE_COLOR)
MySelectedColor = vNewValue
PropertyChanged "SelectedColor"
End Property

Public Property Get SelectedTextColor() As OLE_COLOR
SelectedTextColor = MySelectedTextColor
End Property

Public Property Let SelectedTextColor(ByVal vNewValue As OLE_COLOR)
MySelectedTextColor = vNewValue
PropertyChanged "SelectedTextColor"
End Property

Public Property Get HeaderBackColor() As OLE_COLOR
HeaderBackColor = HeadPic.BackColor
End Property

Public Property Let HeaderBackColor(ByVal vNewValue As OLE_COLOR)
HeadPic.BackColor = vNewValue
DrawHdCap
PropertyChanged "HeaderBackColor"
End Property

Public Property Get HeaderForeColor() As OLE_COLOR
HeaderForeColor = HeadPic.ForeColor
End Property

Public Property Let HeaderForeColor(ByVal vNewValue As OLE_COLOR)
HeadPic.ForeColor = vNewValue
DrawHdCap
PropertyChanged "HeaderForeColor"
End Property

Public Property Get ColumnForeColor() As OLE_COLOR
ColumnForeColor = CapPic.ForeColor
End Property

Public Property Let ColumnForeColor(ByVal vNewValue As OLE_COLOR)
CapPic.ForeColor = vNewValue
DrawColCap
PropertyChanged "ColumnForeColor"
End Property

Public Property Get ListFont() As StdFont
Set ListFont = ListPic.Font
End Property

Public Property Set ListFont(ByVal NewFont As StdFont)
Set ListPic.Font = NewFont
Set TxtEdit.Font = NewFont
RefreshList
PropertyChanged "ListFont"
End Property
Public Property Get ColumnBackColor() As OLE_COLOR
ColumnBackColor = CapPic.BackColor
End Property

Public Property Let ColumnBackColor(ByVal vNewValue As OLE_COLOR)
CapPic.BackColor = vNewValue
For I = 0 To MyColCount
CapSep(I).BackColor = vNewValue
Next I
DrawColCap
PropertyChanged "ColumnBackColor"
End Property

Public Property Get MultiSelTextColor() As OLE_COLOR
MultiSelTextColor = MyMultiSelTextColor
End Property

Public Property Let MultiSelTextColor(ByVal vNewValue As OLE_COLOR)
MyMultiSelTextColor = vNewValue
RefreshList
PropertyChanged "MultiSelTextColor"
End Property

Public Property Get MultiNotSelTextColor() As OLE_COLOR
MultiNotSelTextColor = MyMultiNotSelTextColor
End Property

Public Property Let MultiNotSelTextColor(ByVal vNewValue As OLE_COLOR)
MyMultiNotSelTextColor = vNewValue
RefreshList
PropertyChanged "MultiNotSelTextColor"
End Property

Public Property Get BackPicStyle() As mBackPicStyle
BackPicStyle = MyBackPicStyle
End Property

Public Property Let BackPicStyle(ByVal vNewValue As mBackPicStyle)
MyBackPicStyle = vNewValue
DrawBackPic
PropertyChanged "BackPicStyle"
End Property

Private Sub DrawBackPic()
Dim X1, Y1 As Single
ListPic.Cls
Select Case MyBackPicStyle
Case 0
    If BackPic Then ListPic.PaintPicture BackPic.Picture, 0, 0
    If BackPic Then PicContainer.PaintPicture BackPic.Picture, 0, 0
Case 1
    If BackPic Then
    BackPic.Width = ScaleX(BackPic.Picture.Width, vbHimetric, vbTwips)
    BackPic.Height = ScaleY(BackPic.Picture.Height, vbHimetric, vbTwips)
    For Y1 = 0 To ListPic.ScaleHeight Step BackPic.ScaleHeight
    For X1 = 0 To ListPic.ScaleWidth Step BackPic.ScaleWidth
    ListPic.PaintPicture BackPic.Picture, X1, Y1
    Next X1
    Next Y1
    For Y1 = 0 To PicContainer.ScaleHeight Step BackPic.ScaleHeight
    For X1 = 0 To PicContainer.ScaleWidth Step BackPic.ScaleWidth
    PicContainer.PaintPicture BackPic.Picture, X1, Y1
    Next X1
    Next Y1
    End If
Case 2
    If BackPic Then ListPic.PaintPicture BackPic.Picture, 0, 0, ListPic.Width, ListPic.Height, 0, 0, BackPic.Width, BackPic.Height
    If BackPic Then PicContainer.PaintPicture BackPic.Picture, 0, 0, PicContainer.Width, PicContainer.Height, 0, 0, BackPic.Width, BackPic.Height
End Select
End Sub

Public Property Get HeaderBorder() As mListBorder
HeaderBorder = MyHeaderBorder
End Property

Public Property Let HeaderBorder(ByVal vNewValue As mListBorder)
MyHeaderBorder = vNewValue
DrawHdCap
PropertyChanged "HeaderBorder"
End Property

Public Property Get ColumnBorder() As mListBorder
ColumnBorder = MyColumnBorder
End Property

Public Property Let ColumnBorder(ByVal vNewValue As mListBorder)
MyColumnBorder = vNewValue
DrawColCap
PropertyChanged "ColumnBorder"
End Property

Public Property Get ListBorder() As mListBorder
ListBorder = MyListBorder
End Property

Public Property Let ListBorder(ByVal vNewValue As mListBorder)
MyListBorder = vNewValue
DrawBorder
PropertyChanged "ListBorder"
End Property

Private Sub DrawBorder()
For I = 0 To MyColCount
Make3D CapSep(I), 0.3, mRaised, &HFEFEFE, &HA0A0A0
Next I
Select Case MyHeaderBorder
Case 0

Case 1
Make3D HeadPic, 0.3, mInsert, &HFEFEFE, &HA0A0A0
Case 2
Make3D HeadPic, 0.3, mRaised, &HFEFEFE, &HA0A0A0
Case 3
Make3D HeadPic, 0.8, mInsert, &HFEFEFE, &HA0A0A0
Case 4
Make3D HeadPic, 0.8, mRaised, &HFEFEFE, &HA0A0A0
Case 5
Make3D HeadPic, 0.8, mInsert, &HFEFEFE, &HA0A0A0
Make3D HeadPic, 0.3, mRaised, &HFEFEFE, &HA0A0A0
Case 6
Make3D HeadPic, 0.8, mRaised, &HFEFEFE, &HA0A0A0
Make3D HeadPic, 0.3, mInsert, &HFEFEFE, &HA0A0A0
End Select

Select Case MyColumnBorder
Case 0
For I = 0 To MyColCount
Make3D CapSep(I), 0.3, mInsert, &HFEFEFE, &HA0A0A0
Next I
Case 1
Make3D CapPic, 0.3, mInsert, &HFEFEFE, &HA0A0A0
For I = 0 To MyColCount
Make3D CapSep(I), 0.3, mRaised, &HFEFEFE, &HA0A0A0
Next I
Case 2
Make3D CapPic, 0.3, mRaised, &HFEFEFE, &HA0A0A0
For I = 0 To MyColCount
Make3D CapSep(I), 0.3, mInsert, &HFEFEFE, &HA0A0A0
Next I
Case 3
Make3D CapPic, 0.8, mInsert, &HFEFEFE, &HA0A0A0
For I = 0 To MyColCount
Make3D CapSep(I), 0.8, mRaised, &HFEFEFE, &HA0A0A0
Next I
Case 4
Make3D CapPic, 0.8, mRaised, &HFEFEFE, &HA0A0A0
For I = 0 To MyColCount
Make3D CapSep(I), 0.8, mInsert, &HFEFEFE, &HA0A0A0
Next I
Case 5
Make3D CapPic, 0.8, mInsert, &HFEFEFE, &HA0A0A0
Make3D CapPic, 0.3, mRaised, &HFEFEFE, &HA0A0A0
For I = 0 To MyColCount
Make3D CapSep(I), 0.8, mRaised, &HFEFEFE, &HA0A0A0
Make3D CapSep(I), 0.3, mInsert, &HFEFEFE, &HA0A0A0
Next I
Case 6
Make3D CapPic, 0.8, mInsert, &HFEFEFE, &HA0A0A0
Make3D CapPic, 0.3, mRaised, &HFEFEFE, &HA0A0A0
For I = 0 To MyColCount
Make3D CapSep(I), 0.8, mInsert, &HFEFEFE, &HA0A0A0
Make3D CapSep(I), 0.3, mRaised, &HFEFEFE, &HA0A0A0
Next I
End Select

Select Case MyListBorder
Case 0
DrawBackPic
Case 1
Make3D ListPic, 0.3, mInsert, &HFEFEFE, &HA0A0A0
Case 2
Make3D ListPic, 0.3, mRaised, &HFEFEFE, &HA0A0A0
Case 3
Make3D ListPic, 0.8, mInsert, &HFEFEFE, &HA0A0A0
Case 4
Make3D ListPic, 0.8, mRaised, &HFEFEFE, &HA0A0A0
Case 5
Make3D ListPic, 0.8, mInsert, &HFEFEFE, &HA0A0A0
Make3D ListPic, 0.3, mRaised, &HFEFEFE, &HA0A0A0
Case 6
Make3D ListPic, 0.8, mRaised, &HFEFEFE, &HA0A0A0
Make3D ListPic, 0.3, mInsert, &HFEFEFE, &HA0A0A0
End Select

Select Case MyControlBorder
Case 0

Case 1
Make3D PicBorder, 0.3, mInsert, &HFEFEFE, &HA0A0A0
Case 2
Make3D PicBorder, 0.3, mRaised, &HFEFEFE, &HA0A0A0
Case 3
Make3D PicBorder, 0.8, mInsert, &HFEFEFE, &HA0A0A0
Case 4
Make3D PicBorder, 0.8, mRaised, &HFEFEFE, &HA0A0A0
Case 5
Make3D PicBorder, 0.8, mInsert, &HFEFEFE, &HA0A0A0
Make3D PicBorder, 0.3, mRaised, &HFEFEFE, &HA0A0A0
Case 6
Make3D PicBorder, 0.8, mRaised, &HFEFEFE, &HA0A0A0
Make3D PicBorder, 0.3, mInsert, &HFEFEFE, &HA0A0A0
End Select

End Sub

Public Property Get ControlBorder() As mListBorder
ControlBorder = MyControlBorder
End Property

Public Property Let ControlBorder(ByVal vNewValue As mListBorder)
MyControlBorder = vNewValue
DrawBorder
PropertyChanged "ControlBorder"
End Property

Private Function FindCol(MyX As Single) As Integer
Dim TmpColW As Long
Dim TmpPos As Long

For I = 0 To MyColCount
TmpPos = Split(MyColWidth, "¦")(I)
TmpColW = TmpColW + TmpPos
If MyX < TmpColW Then
FindCol = I
Exit For
End If
Next I
If FindCol >= MyColCount Then
FindCol = MyColCount - 1
End If

End Function

Public Sub SaveTxtEdit()
Dim TmpStr As String

TmpStr = ""
For I = 0 To MyColIndex - 1
TmpStr = TmpStr & Split(MyItem(MyListIndex), MyColSeparator)(I) & MyColSeparator
Next I
TmpStr = TmpStr & TxtEdit.Text & MyColSeparator
For I = MyColIndex + 1 To MyColCount - 1
TmpStr = TmpStr & Split(MyItem(MyListIndex), MyColSeparator)(I) & MyColSeparator
Next I
TmpStr = Left(TmpStr, Len(TmpStr) - 1)
ReplaceItem MyListIndex, TmpStr

End Sub

Public Property Get ColIndex() As Integer
ColIndex = MyColIndex
End Property


Public Property Get EditableList() As Boolean
Attribute EditableList.VB_Description = "You can edit the ListBox directly on it (Press ""Enter"" to save or ""Escape"" to cancel)"
Attribute EditableList.VB_ProcData.VB_Invoke_Property = "General"
EditableList = MyListEdit
End Property

Public Property Let EditableList(ByVal vNewValue As Boolean)
MyListEdit = vNewValue
PropertyChanged "EditableList"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "General"
Enabled = ListPic.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
ListPic.Enabled = vNewValue
PropertyChanged "Enabled"
End Property

Public Property Get IntegralHeight() As Boolean
IntegralHeight = MyIntegralHeight
End Property

Public Property Let IntegralHeight(ByVal vNewValue As Boolean)
MyIntegralHeight = vNewValue
ResizeList
PropertyChanged "IntegralHeight"
End Property

Public Property Get List(Index As Integer) As String
List = MyItem(Index)
End Property


