VERSION 5.00
Begin VB.Form frmConfirm 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "업데이트 정보"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "취소(&C)"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "업데이트 정보"
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "업데이트(&U)"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   2535
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "업데이트를 하시겠습니까?"
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   2880
         Width           =   2130
      End
   End
End
Attribute VB_Name = "frmConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
Unload Me
End Sub
