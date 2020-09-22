VERSION 5.00
Object = "{A34E142F-62D1-42D3-A121-2BAFC6BA2E73}#1.0#0"; "CoolListBox.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin CooListBox.CoolListBox CoolListBox1 
      Height          =   5910
      Left            =   225
      TabIndex        =   8
      Top             =   585
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   10425
      MultiSelect     =   0   'False
      EditableList    =   -1  'True
      ScrollBarAuto   =   2
      ColumnCount     =   5
      ColumnWidth     =   "400¦1800¦1000¦1800¦1000¦1000¦1000¦1000¦1000¦1000¦1000¦1000¦1000¦1000¦1000¦1000¦1000¦1000¦1000¦1000¦"
      ColumnCaption   =   $"Form1.frx":0000
      ColCaptionAlignment=   "0¦1¦2¦2¦1¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦"
      ColListAlignment=   "0¦1¦2¦2¦1¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦0¦"
      BackPicture     =   "Form1.frx":0087
      BackPicStyle    =   1
      TextColor       =   16711808
      SelectedColor   =   8388672
      SelectedTextColor=   12632256
      HeaderBackColor =   14737632
      HeaderForeColor =   8388608
      ColumnBackColor =   12632256
      ColumnForeColor =   0
      MultiNotSelTextColor=   16711935
      ColumnBackColor =   12632256
      ColumnBackColor =   12632256
      ColumnBackColor =   12632256
      ColumnBackColor =   12632256
      ColumnBackColor =   12632256
      ColumnBackColor =   12632256
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ColumnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "MultiSelect"
      Height          =   420
      Left            =   6615
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3285
      Width           =   1770
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ListIndex=ListIndex+1"
      Height          =   420
      Left            =   6615
      TabIndex        =   6
      Top             =   2385
      Width           =   1770
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clear"
      Height          =   420
      Left            =   6615
      TabIndex        =   5
      Top             =   2835
      Width           =   1770
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Replace With Text"
      Height          =   420
      Left            =   6615
      TabIndex        =   4
      Top             =   1035
      Width           =   1770
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remove Sel Item"
      Height          =   420
      Left            =   6615
      TabIndex        =   3
      Top             =   1935
      Width           =   1770
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add Text"
      Height          =   420
      Left            =   6615
      TabIndex        =   2
      Top             =   585
      Width           =   1770
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Text            =   "Text1|Text2|Text3|Text4|Text5"
      Top             =   90
      Width           =   3210
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fill list"
      Height          =   420
      Left            =   6615
      TabIndex        =   0
      Top             =   1485
      Width           =   1770
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
CoolListBox1.MultiSelect = Check1.Value
End Sub

Private Sub Command1_Click()
Dim I

For I = 1 To 20
CoolListBox1.AddItem I & "|Item|Item #:" & I & "|some text here|$" & Format(I * 1.11, "0.00")
Next I
CoolListBox1.HeaderCaption = CoolListBox1.ListCount & " Item(s) in list"
End Sub

Private Sub Command2_Click()
CoolListBox1.AddItem Text1.Text, CoolListBox1.ListIndex

CoolListBox1.HeaderCaption = CoolListBox1.ListCount & " Item(s) in list"
End Sub
Private Sub Command3_Click()
CoolListBox1.RemoveItem (CoolListBox1.ListIndex)

CoolListBox1.HeaderCaption = CoolListBox1.ListCount & " Item(s) in list"
End Sub



Private Sub Command4_Click()
CoolListBox1.ReplaceItem CoolListBox1.ListIndex, Text1.Text
End Sub


Private Sub Command5_Click()
CoolListBox1.Clear
CoolListBox1.HeaderCaption = CoolListBox1.ListCount & " Item(s) in list"
End Sub


Private Sub Command6_Click()
CoolListBox1.ListIndex = CoolListBox1.ListIndex + 1
End Sub

