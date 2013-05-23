VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl FileOpenCtl 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtBody 
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2355
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"FileOpenCtl.ctx":0000
   End
End
Attribute VB_Name = "FileOpenCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Sub UserControl_Initialize()
txtBody.Text = ""
txtBody.FileName = ""
UserControl_Resize
End Sub

Private Sub UserControl_Resize()
On Error GoTo Err_Resize
txtBody.Top = 0
txtBody.Left = 0
With UserControl
txtBody.Width = .Width
txtBody.Height = .Height
End With
Exit Sub
Err_Resize:
Err.Clear
End Sub
