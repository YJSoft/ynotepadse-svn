VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{3DFE7837-69AB-4367-B5DD-159A7EBDE1E6}#15.0#0"; "Mcc.ocx"
Begin VB.Form frmProgressBar 
   BorderStyle     =   1  '���� ����
   Caption         =   "Y's Notepad SE Updater"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5715
   Icon            =   "frmProgressBar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5715
   StartUpPosition =   2  'ȭ�� ���
   Begin RichTextLib.RichTextBox updinfoctl 
      Height          =   2415
      Left            =   120
      TabIndex        =   18
      Top             =   2040
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4260
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmProgressBar.frx":6852
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   255
      Left            =   4440
      TabIndex        =   17
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmProgressBar.frx":68EF
   End
   Begin VB.TextBox txtLocal 
      Height          =   300
      Left            =   480
      TabIndex        =   16
      Text            =   "C:\Version.txt"
      Top             =   5160
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtURL 
      Height          =   300
      Left            =   720
      TabIndex        =   15
      Text            =   "http://ysoftware.dothome.co.kr/Version.txt"
      Top             =   4920
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Frame Frame3 
      Caption         =   "������Ʈ ����(���� ���� �ȵ� ��)�� yyj9411@naver.com����"
      Height          =   1875
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5535
      Begin MyConvenientControl.Mcc Mcc 
         Left            =   4560
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "����"
         Height          =   615
         Left            =   3840
         TabIndex        =   14
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "������Ʈ"
         Enabled         =   0   'False
         Height          =   615
         Left            =   2040
         TabIndex        =   13
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdChk 
         Caption         =   "�ֽ� ���� Ȯ��"
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1815
      End
      Begin ComctlLib.ProgressBar pb 
         Height          =   330
         Left            =   120
         TabIndex        =   8
         Top             =   180
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   582
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.Label lblDownloadSpeed 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�ٿ�ε� �ӵ�: 0 KB/s"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   780
         Width           =   1845
      End
      Begin VB.Label lblDownloadSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�ٿ�ε� ũ��: 0 Bytes / 0 Bytes"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   560
         Width           =   2730
      End
      Begin VB.Label lblPercent 
         Alignment       =   2  '��� ����
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "0.0 %"
         Height          =   180
         Left            =   4920
         TabIndex        =   9
         Top             =   255
         Width           =   450
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " �̾�ޱ� ��� "
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  '����
         Height          =   375
         Left            =   60
         ScaleHeight     =   375
         ScaleWidth      =   1575
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   180
         Width           =   1575
         Begin VB.CheckBox Check1 
            Caption         =   "�̾�ޱ� ���"
            Height          =   255
            Left            =   60
            TabIndex        =   2
            Top             =   60
            Value           =   1  'Ȯ��
            Width           =   1455
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " �ٿ�ε� �ּ� "
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   5295
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  '����
         Height          =   735
         Left            =   60
         ScaleHeight     =   735
         ScaleWidth      =   5055
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   5055
         Begin VB.CommandButton Command2 
            Caption         =   "������?"
            Height          =   300
            Left            =   3840
            TabIndex        =   4
            Top             =   360
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "�ٿ�ε�"
            Height          =   300
            Left            =   3840
            TabIndex        =   3
            Top             =   60
            Visible         =   0   'False
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Base 0

Private Declare Sub RtlMoveMemory Lib "kernel32.dll" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As Long _
)
Private Declare Function VirtualProtect Lib "kernel32.dll" ( _
    ByRef lpAddress As Any, _
    ByVal dwSize As Long, _
    ByVal flNewProtect As Long, _
    ByRef lpflOldProtect As Long _
) As Long
Private Declare Function IsBadReadPtr Lib "kernel32.dll" ( _
    ByRef lp As Any, _
    ByVal ucb As Long _
) As Long
Private Declare Function IsBadWritePtr Lib "kernel32.dll" ( _
    ByRef lp As Any, _
    ByVal ucb As Long _
) As Long
Private Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" ( _
    ByVal lpszUrlName As String _
) As Long
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" ( _
    ByVal hwnd As Long, _
    ByVal szApp As String, _
    ByVal szOtherStuff As String, _
    ByVal hIcon As Long _
) As Long
Private oClsObject As clsIBS
Private oIBSObject As IBindStatusCallback
Private Const PAGE_EXECUTE_READWRITE& = &H40&
Private Const VTBLENTRY_ONPROGRESS& = 6&
Private Const ERROR_ACCESS_DENIED& = 5&
Private Const E_OUTOFMEMORY& = &H8007000E
Private Const E_ABORT& = &H80004004
Private Downloading As Boolean

Private Sub cmdChk_Click()
On Error Resume Next
txtURL.Text = UpdateSite & "Version.txt"
txtLocal.Text = AppPath & "\Version.txt"
Command1_Click
Me.RichTextBox1.FileName = AppPath & "\Version.txt"
txtURL.Text = UpdateSite & "UPDINFO.txt"
txtLocal.Text = AppPath & "UPDINFO.txt"
Command1_Click
updinfoctl.FileName = AppPath & "UPDINFO.txt"
'Text1.Text = updinfoctl.Text
If Not Me.RichTextBox1.Text = GetSetting("YNOTEPADSE", "Program", "Date", "2000-01-01") Then '������ �ٸ���
ForceUpDate:
    Me.cmdChk.Caption = Me.RichTextBox1.Text
    Me.cmdUpdate.Enabled = True
Else
    If LenB(Dir(AppPath & "\Y's Notepad SE.exe")) = 0 Then
        Me.cmdChk.Caption = "������ ����"
        Me.cmdUpdate.Enabled = True '������ �����Ƿ� ������ ������Ʈ
    Else
        Me.cmdChk.Caption = "�ֽ� ����"
        Me.cmdUpdate.Enabled = False
    End If
End If
Mcc.File.DeleteFile AppPath & "\Version.txt" '���� ���� ���� ����
Me.cmdChk.Enabled = 0
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdUpdate_Click()
    cmdUpdate.Enabled = 0
    txtURL.Text = UpdateSite & "YNOTE_" & Me.RichTextBox1.Text & ".exe"
    txtLocal.Text = AppPath & "\YNOTE_" & Me.RichTextBox1.Text & ".exe"
    cmdUpdate.Caption = "�ٿ�ε� �غ���"
    Sleep 1000
    Command1_Click
    cmdUpdate.Caption = "�ٿ�ε� ��"
    On Error Resume Next
    cmdUpdate.Caption = "Y's Notepad SE ����"
    'Shell "tskill Y's Notepad SE.exe"
    KillProcessByName "Y's Notepad SE.exe"
    Sleep 1000
    cmdUpdate.Caption = "���� �����"
    Sleep 2000
    'Mcc.Registry.WriteRegistry  '������Ʈ�� ���� ���μ��� ���� ����
    'Kill AppPath & "\Y's Notepad SE.exe" '������Ʈ �� ������ ���� ����
    Mcc.File.RenameFile AppPath & "\Y's Notepad SE.exe", AppPath & "\Y's Notepad SE.exe.old" '������Ʈ �� ���� ���(���� ���ɼ�)
    On Error GoTo Err_Update
    cmdUpdate.Caption = "��ġ��"
    Sleep 2000
    Mcc.File.RenameFile AppPath & "\YNOTE_" & Me.RichTextBox1.Text & ".exe", AppPath & "\Y's Notepad SE.exe" '������Ʈ
    cmdUpdate.Caption = "������"
    Sleep 1000
    Mcc.File.DeleteFile AppPath & "\Y's Notepad SE.exe.old" '��� ���� ����
    SaveSetting "YNOTEPADSE", "Program", "Date", Me.RichTextBox1.Text
    'Me.RichTextBox1.FileName = "" '��ġ�ع� ���� �ε� ����
    'Mcc.File.DeleteFile AppPath & "\Version.txt" '���� ���� ���� ����
    MsgBox "������Ʈ�� �Ϸ�Ǿ����ϴ�.", vbInformation, "������Ʈ ����"
    End
    Exit Sub
Err_Update:
MsgBox "������Ʈ �� ���� �߻�!" & vbCr & "������ȣ:" & Err.Number & vbCr & Err.Description, vbCritical
Mcc.File.RenameFile AppPath & "\Y's Notepad SE.exe.old", AppPath & "\Y's Notepad SE.exe" '������Ʈ ���з� ���� ����
Err.Clear
End '������Ʈ ������ ����
End Sub

Private Sub Command1_Click()
    Dim ret As Long
    If Not Downloading Then
        ' �̾�ޱ�(ĳ��) ����� ���ϴ� ��쿡��...
        If 1 Then '�̾�ޱ� ��Ȱ��ȭ
            ' ĳ�� ��Ʈ���� �����ش�.
            If DeleteUrlCacheEntry(txtURL.Text) = ERROR_ACCESS_DENIED Then
                ' ������ �ź� ������ ��ȯ�� ���
                MsgBox "������ ĳ�ø� ������ �� �������ϴ�. " & _
                       "������ �ֱٿ� ������Ʈ �� ��쿡�� ������ ������ ���� �� �ֽ��ϴ�.", vbExclamation, "���"
            End If
        End If
        
        SpeedCheck = False
        Cancelled = False
        Downloading = True
        Command1.Caption = "���"
        txtURL.Enabled = False
        txtLocal.Enabled = False
        
        ' �ٿ�ε带 �����Ѵ�.
        ret = olelib.URLDownloadToFile(Nothing, txtURL.Text, txtLocal.Text, 0&, oIBSObject)
        If ExitOnCancel Then
            End
        End If
        
        ' ��ȯ���� üũ�Ͽ� �ٿ�ε� ���������� �ʾҴ��� üũ�Ͽ���.
        ' (����: �ٿ�ε� ������ S_OK ��ȯ)
        If ret = E_OUTOFMEMORY Then
            MsgBox "������ �ٿ�ε��� �� �����ϴ�. �޸𸮰� �����մϴ�.", vbCritical, "����"
        ElseIf ret = E_ABORT Then
            ' ��ҵ� ���
            MsgBox "�ٿ�ε尡 �ߴܵǾ����ϴ�.", vbExclamation, "�ٿ�ε�"
        ElseIf ret Then
            'MsgBox "URLDownloadToFile Failed because of unknown reason." & vbCrLf & vbCrLf & _
                   "URLDownloadToFile() Error Code: " & ret & " (0x" & Hex$(ret) & ")", vbCritical, "API Error"
            MsgBox "�� �� ���� ������ ������Ʈ ������ ������ �� �������ϴ�." & vbCr & "���ͳ� ������ ���������� Ȯ���Ͽ� �ֽʽÿ�." & vbCr & "���� �ڵ�:" & ret & " (0x" & Hex$(ret) & ")", vbCritical, "������Ʈ ����!"
            If Not ErrSkip Then End
        Else
            ' �ٿ�ε�� �Ϸ��ߴµ� ������ ������ ���� ���
            If LenB(Dir(txtLocal.Text)) = 0 Then
                MsgBox "������ �ٿ�ε��� �� �����ϴ�. ������ ������ �� �������ϴ�." & vbCr & "������ ������̰ų� �б� �����Դϴ�", vbCritical, "����"
            End If
        End If
        
        Downloading = False
        Command1.Caption = "�ٿ�ε�"
        Command1.Enabled = True
        txtURL.Enabled = True
        txtLocal.Enabled = True
    Else
        Downloading = False
        Cancelled = True
        Command1.Enabled = False
    End If
End Sub



Private Sub Command5_Click()

End Sub

Private Sub Form_Load()
    '������Ʈ ����Ʈ ������ �ҷ��´�
    Me.RichTextBox1.FileName = AppPath & "\UpdateSite.ini"
    UpdateSite = Me.RichTextBox1.Text
    Me.RichTextBox1.Text = ""
    Me.RichTextBox1.FileName = ""
    '�������� ��ü ������Ʈ ����
    txtURL.Text = UpdateSite & "UDVersion.txt"
    txtLocal.Text = AppPath & "\UDVersion.txt"
    Command1_Click
    Me.RichTextBox1.FileName = AppPath & "\UDVersion.txt"
    If Not Me.RichTextBox1.Text = "2013-04-03" Then '������ �ٸ���(������Ʈ �ʿ�)
        MsgBox "���ο� ������ �ֽ��ϴ�. �� ������ �޾� �ɴϴ�...", vbInformation, "������Ʈ Ȯ�ε�"
        txtURL.Text = UpdateSite & "UPDATER_" & Me.RichTextBox1.Text & ".exe"
        txtLocal.Text = AppPath & "\UPDATER_" & Me.RichTextBox1.Text & ".exe"
        Command1_Click
        On Error Resume Next
        Sleep 1000
        Mcc.File.RenameFile AppPath & "\Updater.exe", AppPath & "\Updater.exe.old" '������Ʈ �� ���� ���(���� ���ɼ�)
        Sleep 2000
        Mcc.File.RenameFile AppPath & "\UPDATER_" & Me.RichTextBox1.Text & ".exe", AppPath & "\Updater.exe" '������Ʈ
        Sleep 1000
        Mcc.File.DeleteFile AppPath & "\Updater.exe.old" '��� ���� ����
        MsgBox "���������� ������Ʈ�� �Ϸ�Ǿ����ϴ�." & "�������͸� ������� �ּ���.", vbInformation, "������Ʈ ����"
        End
    End If
    Dim VTableBase As Long, pProgressFuncPtr As Long, blChangedProtect As Boolean, lngOldProtect As Long
    Mcc.DisableXButton Me.hwnd
    If LenB(Dir(AppPath & "\UpdateSite.ini")) = 0 Then
    'AppPath & "\Y's Notepad SE.exe"
        Me.Hide
        MsgBox "������Ʈ ���� ���� ������ ã�� �� �����ϴ�!", vbCritical, "ġ������ ����"
        End
    End If
    
    'If App.LogMode = 0 Then
    '    If MsgBox("�� ���������� ������ ��� ���� Ŭ���� ����� �Լ� �����͸� �����մϴ�." & vbCrLf & vbCrLf & _
    '              "Visual Basic IDE ��忡�� ������ ��쿡�� �ټ� �Ҿ����� �� �ֽ��ϴ�." & vbCrLf & vbCrLf & _
    '              "��� �����Ͻðڽ��ϱ�?", vbExclamation Or vbYesNo, "���") = vbNo Then
    '        End
    '    End If
    'End If
    
    ' �۾� �����ڵ�� ������ �̸�
    App.Title = Caption
    
    ' Ŭ���� ����
    Set oClsObject = New clsIBS
    Set oIBSObject = oClsObject
    
    ' IBindStatusCallback::Virtual Table �����͸� ����
    RtlMoveMemory VTableBase, ByVal ObjPtr(oIBSObject), 4&
    If VTableBase = 0& Then
        MsgBox "IBindStatusCallback::Virtual Table �ּҰ� �߸��Ǿ����ϴ�.", vbCritical, "����"
        End
    ElseIf IsBadReadPtr(ByVal VTableBase, 4&) Then
        MsgBox "IBindStatusCallback::Virtual Table �ּҰ� �߸��Ǿ����ϴ�.", vbCritical, "����"
        End
    End If
    
    ' VTable�κ��� OnProgress() �Լ� �����Ͱ� ����ִ� �޸��� �ּҸ� ����.
    pProgressFuncPtr = VTableBase + 4& * VTBLENTRY_ONPROGRESS
    ' �޸� ��ȣ �ɼ��� �˻��Ͽ� �������Ѱ�� ���Ƿ� ����
    If IsBadWritePtr(ByVal pProgressFuncPtr, 4&) Then
        blChangedProtect = True
        If VirtualProtect(ByVal pProgressFuncPtr, _
                          4&, _
                          PAGE_EXECUTE_READWRITE, _
                          lngOldProtect) = 0& Then
            MsgBox "VirtualProtect() �Լ��� �����߽��ϴ�." & vbCrLf & vbCrLf & _
                   "��ȯ�� �����ڵ�� " & Err.LastDllError & " �Դϴ�.", vbCritical, "API ����"
            End
        End If
    End If
    
    ' IBindStatusCallback::OnProgress() �Լ��� ��� �� �Լ��� Redirect ��ŵ�ϴ�.
    ' Reason: VB�� Ŭ���� ����� �״�� �Ѱ��൵ ������ ��ȯ���� S_OK�� �׻� ��ȯ�Ǳ� ������
    ' ������ ��� ��� �Ұ����ϴٴ� ������ �־� �̷��� ����.
    RtlMoveMemory ByVal pProgressFuncPtr, Conversion.CLng(AddressOf RealOnProgress), 4&
    
    ' �޸� ��ȣ �ɼ��� ���۵� ��� �ٽ� ������� ����
    If blChangedProtect Then
        VirtualProtect ByVal pProgressFuncPtr, _
                       4&, _
                       lngOldProtect, _
                       0&
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Downloading Then
        If MsgBox("�ٿ�ε尡 �������Դϴ�. �׷��� �����Ͻðڽ��ϱ�?", vbExclamation Or vbYesNo, "�ȳ�") = vbYes Then
            App.TaskVisible = False
            If Downloading Then
                Me.Hide
                Downloading = False
                Cancelled = True
                ExitOnCancel = True
                Cancel = 1
            End If
        Else
            Cancel = 1
        End If
    End If
End Sub

Private Sub RichTextBox2_Change()

End Sub

