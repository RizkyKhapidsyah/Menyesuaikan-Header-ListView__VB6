VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Menyesuaikan Header ListView"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   975
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1720
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  With ListView1
     .ColumnHeaders.Add , , "Header 1"
     .ColumnHeaders.Add , , "Header 2"
     .ColumnHeaders.Add , , "Header 3"
  End With
End Sub
  
Private Sub Command1_Click()
  Dim Column As Long
  Dim Counter As Long
  Counter = 0
  For Column = Counter To _
               ListView1.ColumnHeaders.Count - 2
     SendMessage ListView1.hWnd, LVM_SETCOLUMNWIDTH, _
                 Column, LVSCW_AUTOSIZE_USEHEADER
  Next
End Sub


