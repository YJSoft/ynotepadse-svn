Attribute VB_Name = "modProgress"

Option Explicit
Option Base 0

Private Declare Function GetTickCount _
    Lib "kernel32.dll" () As Long
Private Const S_OK& = 0&
Private Const E_ABORT& = &H80004004

Public Cancelled As Boolean, SpeedCheck As Boolean, ExitOnCancel As Boolean
Private SpeedLastTick As Long, SpeedLastSize As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Private Function ExpressByte(ByVal lFileLen As Long) As String
    If lFileLen > 1073741824 Then
        ExpressByte = Round(lFileLen / 1073741824, 1) & " GB"
    ElseIf lFileLen > 1048576 Then
        ExpressByte = Round(lFileLen / 1048576, 1) & " MB"
    ElseIf lFileLen > 1024& Then
        ExpressByte = Round(lFileLen / 1024, 1) & " KB"
    Else
        ExpressByte = lFileLen & " Bytes"
    End If
End Function

Public Function RealOnProgress(ByVal objIBSMe As IBindStatusCallback, _
                               ByVal ulProgress As Long, _
                               ByVal ulProgressMax As Long, _
                               ByVal ulStatusCode As olelib.BINDSTATUS, _
                               ByVal szStatusText As Long) As Long
    If Cancelled Then
        SpeedCheck = False
        Cancelled = False
        frmProgressBar.Command1.Enabled = True
        frmProgressBar.txtURL.Enabled = True
        frmProgressBar.txtLocal.Enabled = True

        RealOnProgress = E_ABORT
        Exit Function
    Else
        If Not SpeedCheck Then
            SpeedCheck = True
            SpeedLastTick = GetTickCount
            SpeedLastSize = ulProgress
            frmProgressBar.lblDownloadSpeed.Caption = "다운로드 속도: 0 KB/s"
        Else
            If (GetTickCount - SpeedLastTick) > 500& Then
                frmProgressBar.lblDownloadSpeed.Caption = "다운로드 속도: " & ExpressByte((ulProgress - SpeedLastSize) / ((GetTickCount - SpeedLastTick) / 1000)) & "/s"
                SpeedLastTick = GetTickCount
                SpeedLastSize = ulProgress
            End If
        End If
        
        frmProgressBar.lblDownloadSize.Caption = "다운로드 크기: " & ExpressByte(ulProgress) & " / " & ExpressByte(ulProgressMax)
        
        With frmProgressBar.pb
            If .Value Then 'ProgressBar의 값이 0 이상이면
                .Value = 0 '0으로 초기화
            End If
            If ulProgressMax > 0 Then
                .Max = ulProgressMax '최대값
            Else
                .Max = 1 '최대 1
            End If
            If ulProgress > 0 Then
                .Value = ulProgress '현재 다운 완료된 만큼 Progress 표시
            Else
                .Value = 0
            End If
            
            If ulProgressMax > 0 Then
                frmProgressBar.lblPercent.Caption = Format(Round(ulProgress / ulProgressMax * 100, 1), "##0.0") & " %"
            Else
                frmProgressBar.lblPercent.Caption = "N/A"
            End If
        End With
        
        DoEvents
        RealOnProgress = S_OK
    End If
End Function
Public Function AppPath() As String
If Right(App.Path, 1) = "\" Then
    AppPath = Left(App.Path, Len(App.Path) - 1)
Else
    AppPath = App.Path
    Debug.Print App.Path
End If
End Function
