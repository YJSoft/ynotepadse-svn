VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{3DFE7837-69AB-4367-B5DD-159A7EBDE1E6}#15.0#0"; "Mcc.ocx"
Begin VB.Form frmProgressBar 
   BorderStyle     =   1  '단일 고정
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
   StartUpPosition =   2  '화면 가운데
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
      Caption         =   "업데이트 오류(파일 실행 안됨 등)시 yyj9411@naver.com으로"
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
         Caption         =   "종료"
         Height          =   615
         Left            =   3840
         TabIndex        =   14
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "업데이트"
         Enabled         =   0   'False
         Height          =   615
         Left            =   2040
         TabIndex        =   13
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdChk 
         Caption         =   "최신 버전 확인"
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
         BackStyle       =   0  '투명
         Caption         =   "다운로드 속도: 0 KB/s"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   780
         Width           =   1845
      End
      Begin VB.Label lblDownloadSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "다운로드 크기: 0 Bytes / 0 Bytes"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   560
         Width           =   2730
      End
      Begin VB.Label lblPercent 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "0.0 %"
         Height          =   180
         Left            =   4920
         TabIndex        =   9
         Top             =   255
         Width           =   450
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " 이어받기 사용 "
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  '없음
         Height          =   375
         Left            =   60
         ScaleHeight     =   375
         ScaleWidth      =   1575
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   180
         Width           =   1575
         Begin VB.CheckBox Check1 
            Caption         =   "이어받기 사용"
            Height          =   255
            Left            =   60
            TabIndex        =   2
            Top             =   60
            Value           =   1  '확인
            Width           =   1455
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " 다운로드 주소 "
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   5295
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  '없음
         Height          =   735
         Left            =   60
         ScaleHeight     =   735
         ScaleWidth      =   5055
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   5055
         Begin VB.CommandButton Command2 
            Caption         =   "제작자?"
            Height          =   300
            Left            =   3840
            TabIndex        =   4
            Top             =   360
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "다운로드"
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
If Not Me.RichTextBox1.Text = GetSetting("YNOTEPADSE", "Program", "Date", "2000-01-01") Then '버전이 다르다
ForceUpDate:
    Me.cmdChk.Caption = Me.RichTextBox1.Text
    Me.cmdUpdate.Enabled = True
Else
    If LenB(Dir(AppPath & "\Y's Notepad SE.exe")) = 0 Then
        Me.cmdChk.Caption = "파일이 없음"
        Me.cmdUpdate.Enabled = True '파일이 없으므로 강제로 업데이트
    Else
        Me.cmdChk.Caption = "최신 버전"
        Me.cmdUpdate.Enabled = False
    End If
End If
Mcc.File.DeleteFile AppPath & "\Version.txt" '버전 정보 파일 삭제
Me.cmdChk.Enabled = 0
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdUpdate_Click()
    cmdUpdate.Enabled = 0
    txtURL.Text = UpdateSite & "YNOTE_" & Me.RichTextBox1.Text & ".exe"
    txtLocal.Text = AppPath & "\YNOTE_" & Me.RichTextBox1.Text & ".exe"
    cmdUpdate.Caption = "다운로드 준비중"
    Sleep 1000
    Command1_Click
    cmdUpdate.Caption = "다운로드 중"
    On Error Resume Next
    cmdUpdate.Caption = "Y's Notepad SE 종료"
    'Shell "tskill Y's Notepad SE.exe"
    KillProcessByName "Y's Notepad SE.exe"
    Sleep 1000
    cmdUpdate.Caption = "파일 백업중"
    Sleep 2000
    'Mcc.Registry.WriteRegistry  '업데이트를 위해 프로세스 강제 종료
    'Kill AppPath & "\Y's Notepad SE.exe" '업데이트 전 구버전 파일 삭제
    Mcc.File.RenameFile AppPath & "\Y's Notepad SE.exe", AppPath & "\Y's Notepad SE.exe.old" '업데이트 전 파일 백업(오류 가능성)
    On Error GoTo Err_Update
    cmdUpdate.Caption = "패치중"
    Sleep 2000
    Mcc.File.RenameFile AppPath & "\YNOTE_" & Me.RichTextBox1.Text & ".exe", AppPath & "\Y's Notepad SE.exe" '업데이트
    cmdUpdate.Caption = "정리중"
    Sleep 1000
    Mcc.File.DeleteFile AppPath & "\Y's Notepad SE.exe.old" '백업 파일 삭제
    SaveSetting "YNOTEPADSE", "Program", "Date", Me.RichTextBox1.Text
    'Me.RichTextBox1.FileName = "" '리치텍박 파일 로드 해제
    'Mcc.File.DeleteFile AppPath & "\Version.txt" '버전 정보 파일 삭제
    MsgBox "업데이트가 완료되었습니다.", vbInformation, "업데이트 성공"
    End
    Exit Sub
Err_Update:
MsgBox "업데이트 중 오류 발생!" & vbCr & "오류번호:" & Err.Number & vbCr & Err.Description, vbCritical
Mcc.File.RenameFile AppPath & "\Y's Notepad SE.exe.old", AppPath & "\Y's Notepad SE.exe" '업데이트 실패로 파일 복원
Err.Clear
End '업데이트 실패후 종료
End Sub

Private Sub Command1_Click()
    Dim ret As Long
    If Not Downloading Then
        ' 이어받기(캐시) 사용을 안하는 경우에는...
        If 1 Then '이어받기 비활성화
            ' 캐시 엔트리를 지워준다.
            If DeleteUrlCacheEntry(txtURL.Text) = ERROR_ACCESS_DENIED Then
                ' 엑세스 거부 오류가 반환된 경우
                MsgBox "브라우저 캐시를 삭제할 수 없었습니다. " & _
                       "파일이 최근에 업데이트 된 경우에도 이전의 정보를 받을 수 있습니다.", vbExclamation, "경고"
            End If
        End If
        
        SpeedCheck = False
        Cancelled = False
        Downloading = True
        Command1.Caption = "취소"
        txtURL.Enabled = False
        txtLocal.Enabled = False
        
        ' 다운로드를 시작한다.
        ret = olelib.URLDownloadToFile(Nothing, txtURL.Text, txtLocal.Text, 0&, oIBSObject)
        If ExitOnCancel Then
            End
        End If
        
        ' 반환값을 체크하여 다운로드 실패하지는 않았는지 체크하여줌.
        ' (참고: 다운로드 성공시 S_OK 반환)
        If ret = E_OUTOFMEMORY Then
            MsgBox "파일을 다운로드할 수 없습니다. 메모리가 부족합니다.", vbCritical, "오류"
        ElseIf ret = E_ABORT Then
            ' 취소된 경우
            MsgBox "다운로드가 중단되었습니다.", vbExclamation, "다운로드"
        ElseIf ret Then
            'MsgBox "URLDownloadToFile Failed because of unknown reason." & vbCrLf & vbCrLf & _
                   "URLDownloadToFile() Error Code: " & ret & " (0x" & Hex$(ret) & ")", vbCritical, "API Error"
            MsgBox "알 수 없는 이유로 업데이트 서버에 연결할 수 없었습니다." & vbCr & "인터넷 연결이 정상적인지 확인하여 주십시오." & vbCr & "오류 코드:" & ret & " (0x" & Hex$(ret) & ")", vbCritical, "업데이트 오류!"
            If Not ErrSkip Then End
        Else
            ' 다운로드는 완료했는데 파일이 생기지 않은 경우
            If LenB(Dir(txtLocal.Text)) = 0 Then
                MsgBox "파일을 다운로드할 수 없습니다. 파일을 생성할 수 없었습니다." & vbCr & "파일이 사용중이거나 읽기 전용입니다", vbCritical, "오류"
            End If
        End If
        
        Downloading = False
        Command1.Caption = "다운로드"
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
    '업데이트 사이트 정보를 불러온다
    Me.RichTextBox1.FileName = AppPath & "\UpdateSite.ini"
    UpdateSite = Me.RichTextBox1.Text
    Me.RichTextBox1.Text = ""
    Me.RichTextBox1.FileName = ""
    '업데이터 자체 업데이트 조사
    txtURL.Text = UpdateSite & "UDVersion.txt"
    txtLocal.Text = AppPath & "\UDVersion.txt"
    Command1_Click
    Me.RichTextBox1.FileName = AppPath & "\UDVersion.txt"
    If Not Me.RichTextBox1.Text = "2013-04-03" Then '버전이 다르다(업데이트 필요)
        MsgBox "새로운 버전이 있습니다. 새 버전을 받아 옵니다...", vbInformation, "업데이트 확인됨"
        txtURL.Text = UpdateSite & "UPDATER_" & Me.RichTextBox1.Text & ".exe"
        txtLocal.Text = AppPath & "\UPDATER_" & Me.RichTextBox1.Text & ".exe"
        Command1_Click
        On Error Resume Next
        Sleep 1000
        Mcc.File.RenameFile AppPath & "\Updater.exe", AppPath & "\Updater.exe.old" '업데이트 전 파일 백업(오류 가능성)
        Sleep 2000
        Mcc.File.RenameFile AppPath & "\UPDATER_" & Me.RichTextBox1.Text & ".exe", AppPath & "\Updater.exe" '업데이트
        Sleep 1000
        Mcc.File.DeleteFile AppPath & "\Updater.exe.old" '백업 파일 삭제
        MsgBox "업데이터의 업데이트가 완료되었습니다." & "업데이터를 재실행해 주세요.", vbInformation, "업데이트 성공"
        End
    End If
    Dim VTableBase As Long, pProgressFuncPtr As Long, blChangedProtect As Boolean, lngOldProtect As Long
    Mcc.DisableXButton Me.hwnd
    If LenB(Dir(AppPath & "\UpdateSite.ini")) = 0 Then
    'AppPath & "\Y's Notepad SE.exe"
        Me.Hide
        MsgBox "업데이트 서버 설정 파일을 찾을 수 없습니다!", vbCritical, "치명적인 오류"
        End
    End If
    
    'If App.LogMode = 0 Then
    '    If MsgBox("이 예제에서는 동적인 제어를 위해 클래스 모듈의 함수 포인터를 변경합니다." & vbCrLf & vbCrLf & _
    '              "Visual Basic IDE 모드에서 실행한 경우에는 다소 불안정할 수 있습니다." & vbCrLf & vbCrLf & _
    '              "계속 실행하시겠습니까?", vbExclamation Or vbYesNo, "경고") = vbNo Then
    '        End
    '    End If
    'End If
    
    ' 작업 관리자등에서 보여질 이름
    App.Title = Caption
    
    ' 클래스 생성
    Set oClsObject = New clsIBS
    Set oIBSObject = oClsObject
    
    ' IBindStatusCallback::Virtual Table 포인터를 얻음
    RtlMoveMemory VTableBase, ByVal ObjPtr(oIBSObject), 4&
    If VTableBase = 0& Then
        MsgBox "IBindStatusCallback::Virtual Table 주소가 잘못되었습니다.", vbCritical, "오류"
        End
    ElseIf IsBadReadPtr(ByVal VTableBase, 4&) Then
        MsgBox "IBindStatusCallback::Virtual Table 주소가 잘못되었습니다.", vbCritical, "오류"
        End
    End If
    
    ' VTable로부터 OnProgress() 함수 포인터가 들어있는 메모리의 주소를 구함.
    pProgressFuncPtr = VTableBase + 4& * VTBLENTRY_ONPROGRESS
    ' 메모리 보호 옵션을 검사하여 부적절한경우 임의로 조절
    If IsBadWritePtr(ByVal pProgressFuncPtr, 4&) Then
        blChangedProtect = True
        If VirtualProtect(ByVal pProgressFuncPtr, _
                          4&, _
                          PAGE_EXECUTE_READWRITE, _
                          lngOldProtect) = 0& Then
            MsgBox "VirtualProtect() 함수가 실패했습니다." & vbCrLf & vbCrLf & _
                   "반환된 오류코드는 " & Err.LastDllError & " 입니다.", vbCritical, "API 오류"
            End
        End If
    End If
    
    ' IBindStatusCallback::OnProgress() 함수를 모듈 안 함수로 Redirect 시킵니다.
    ' Reason: VB의 클래스 모듈을 그대로 넘겨줘도 되지만 반환값이 S_OK로 항상 반환되기 때문에
    ' 임의의 취소 제어가 불가능하다는 단점이 있어 이렇게 했음.
    RtlMoveMemory ByVal pProgressFuncPtr, Conversion.CLng(AddressOf RealOnProgress), 4&
    
    ' 메모리 보호 옵션이 조작된 경우 다시 원래대로 복구
    If blChangedProtect Then
        VirtualProtect ByVal pProgressFuncPtr, _
                       4&, _
                       lngOldProtect, _
                       0&
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Downloading Then
        If MsgBox("다운로드가 진행중입니다. 그래도 종료하시겠습니까?", vbExclamation Or vbYesNo, "안내") = vbYes Then
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

