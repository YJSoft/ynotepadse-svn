VERSION 5.00
Object = "*\AFileOpenCtl_YJSoft.vbp"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin FileOpenCtl_YJSoft.FileOpenCtl FileOpenCtl1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4471
      FontSize        =   9.75
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If Not FileOpenCtl1.openFile("C:\aaa1.txt") Then MsgBox FileOpenCtl1.ErrNumber & " / " & FileOpenCtl1.ErrDesc
End Sub

Private Sub Form_Resize()
With FileOpenCtl1
.Top = 0
.Left = 0
.Width = Me.ScaleWidth
.Height = Me.ScaleHeight
End With
End Sub
