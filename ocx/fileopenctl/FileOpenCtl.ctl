VERSION 5.00
Begin VB.UserControl FileOpenCtl 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PropertyPages   =   "FileOpenCtl.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox txtBody 
      Height          =   1575
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   0
      Text            =   "FileOpenCtl.ctx":0012
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "FileOpenCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'기본 속성 값:
Const m_def_ErrNumber = 0
Const m_def_ErrDesc = "오류가 없습니다."
'Const m_def_ErrNumber = 0
'Const m_def_ErrDesc = ""
Const m_def_Dirty = 0
Const m_def_FillColor = &H0&
Const m_def_FileName = "0"
'속성 변수:
Dim m_ErrNumber As Integer
Dim m_ErrDesc As String
'Dim m_ErrNumber As Integer
'Dim m_ErrDesc As String
Dim m_Dirty As Boolean
Dim m_FillColor As OLE_COLOR
Dim m_FileName As String
'이벤트 선언:
Event Click() 'MappingInfo=txtBody,txtBody,-1,Click
Attribute Click.VB_Description = "개체에서 마우스 단추를 눌렀다가 놓을 때 발생합니다."
Event DblClick() 'MappingInfo=txtBody,txtBody,-1,DblClick
Attribute DblClick.VB_Description = "마우스 단추를 개체에서 누르고 놓은 후 다시 누르고 놓으면 발생합니다."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtBody,txtBody,-1,KeyDown
Attribute KeyDown.VB_Description = "개체에 포커스가 있을 때 키를 누르면 발생합니다."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtBody,txtBody,-1,KeyPress
Attribute KeyPress.VB_Description = "ANSI키를 누르고 놓았을 경우 발생합니다."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtBody,txtBody,-1,KeyUp
Attribute KeyUp.VB_Description = "개체에 포커스가 있을 때 키를 놓으면 발생합니다."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtBody,txtBody,-1,MouseDown
Attribute MouseDown.VB_Description = "개체에 포커스가 있을 때 마우스 단추를 누르면 발생합니다."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtBody,txtBody,-1,MouseMove
Attribute MouseMove.VB_Description = "마우스를 움직일 경우 발생합니다."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtBody,txtBody,-1,MouseUp
Attribute MouseUp.VB_Description = "개체에 포커스가 있을 때 마우스 단추를 놓으면 발생합니다."
Event Changed() 'MappingInfo=txtBody,txtBody,-1,Change
Attribute Changed.VB_Description = "컨트롤의 내용이 변경될 때 발생합니다."



Private Sub UserControl_Initialize()
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
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=txtBody,txtBody,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "개체에서 텍스트나 그래픽을 표시하는 전경색을 반환하거나 설정합니다."
    ForeColor = txtBody.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtBody.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=txtBody,txtBody,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "사용자가 만든 이벤트에 대해 개체가 응답할 수 있는지의 여부를 결정하는 값을 반환하거나 설정합니다."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "DefaultPage"
    Enabled = txtBody.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtBody.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Private Sub txtBody_Click()
    RaiseEvent Click
End Sub

Private Sub txtBody_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub txtBody_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtBody_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtBody_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtBody_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub txtBody_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txtBody_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=10,3,2,&H0&
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "도형, 원, 상자를 채우는 데 사용된 색을 반환하거나 설정합니다."
Attribute FillColor.VB_MemberFlags = "400"
    If Ambient.UserMode Then Err.Raise 393
    FillColor = m_FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_FillColor = New_FillColor
    PropertyChanged "FillColor"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=txtBody,txtBody,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "개체의 텍스트나 그래픽을 표시하기 위해 사용되는 배경색을 반환하거나 설정합니다."
    BackColor = txtBody.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtBody.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=txtBody,txtBody,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "굵게 글꼴 유형을 반환하거나 설정합니다."
Attribute FontBold.VB_ProcData.VB_Invoke_Property = "DefaultPage"
    FontBold = txtBody.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    txtBody.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=txtBody,txtBody,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "기울임 글꼴 유형을 반환하거나 설정합니다."
Attribute FontItalic.VB_ProcData.VB_Invoke_Property = "DefaultPage"
    FontItalic = txtBody.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    txtBody.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=txtBody,txtBody,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "주어진 단계의 각 행에 나타나는 글꼴의 이름을 지정합니다."
Attribute FontName.VB_ProcData.VB_Invoke_Property = "DefaultPage"
    FontName = txtBody.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    txtBody.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=txtBody,txtBody,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "주어진 단계의 각 행에 나타나는 글꼴 크기를 포인트 단위로 지정합니다."
Attribute FontSize.VB_ProcData.VB_Invoke_Property = "DefaultPage"
    FontSize = txtBody.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    txtBody.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=txtBody,txtBody,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "취소선 글꼴 유형을 반환하거나 설정합니다."
Attribute FontStrikethru.VB_ProcData.VB_Invoke_Property = "DefaultPage"
    FontStrikethru = txtBody.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    txtBody.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=txtBody,txtBody,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "밑줄 글꼴 유형을 반환하거나 설정합니다."
Attribute FontUnderline.VB_ProcData.VB_Invoke_Property = "DefaultPage"
    FontUnderline = txtBody.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    txtBody.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=0
Public Function OpenFile(strFileName As String) As Boolean
Attribute OpenFile.VB_Description = "파일을 엽니다.(오류 발생시 False값이 반환됩니다)"
On Error GoTo err_open
m_FileName = strFileName
    Dim FreeFileNum As Integer
    FreeFileNum = FreeFile
    Open strFileName For Input As #FreeFileNum
    txtBody.Text = StrConv(InputB(LOF(FreeFileNum), FreeFileNum), vbUnicode)
    OpenFile = True
    PropertyChanged "Text"
    Exit Function
err_open:
OpenFile = False
m_ErrNumber = Err.Number
m_ErrDesc = Err.Description
Err.Clear
End Function

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=0
Public Function SaveFile(strFileName As String) As Boolean
Attribute SaveFile.VB_Description = "파일을 저장합니다.(오류 발생시 False값이 반환됩니다)"
On Error GoTo err_save
m_FileName = strFileName
    Dim FreeFileNum As Integer
    FreeFileNum = FreeFile
    Open strFileName For Output As #FreeFileNum
    Print #FreeFileNum, txtBody.Text
    Close #FreeFileNum
    SaveFile = True
    Exit Function
err_save:
SaveFile = False
m_ErrNumber = Err.Number
m_ErrDesc = Err.Description
Err.Clear
End Function

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,1,2,0
Public Property Get FileName() As String
Attribute FileName.VB_Description = "현재 열려진 파일 이름을 반환합니다."
Attribute FileName.VB_MemberFlags = "400"
    FileName = m_FileName
End Property

Public Property Let FileName(ByVal New_FileName As String)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    Err.Raise 387
End Property

Private Sub txtBody_Change()
    RaiseEvent Changed
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=txtBody,txtBody,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "컨트롤에 포함된 텍스트를 반환하거나 설정합니다."
    Text = txtBody.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtBody.Text() = New_Text
    PropertyChanged "Text"
End Property

'사용자 정의 컨트롤에 대한 속성을 초기화합니다.
Private Sub UserControl_InitProperties()
    m_FillColor = m_def_FillColor
    m_FileName = m_def_FileName
    m_Dirty = m_def_Dirty
'    m_ErrNumber = m_def_ErrNumber
'    m_ErrDesc = m_def_ErrDesc
    m_ErrNumber = m_def_ErrNumber
    m_ErrDesc = m_def_ErrDesc
End Sub

'저장소에서 속성값을 로드합니다.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    txtBody.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtBody.Enabled = PropBag.ReadProperty("Enabled", True)
    m_FillColor = PropBag.ReadProperty("FillColor", m_def_FillColor)
    txtBody.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtBody.FontBold = PropBag.ReadProperty("FontBold", 0)
    txtBody.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    txtBody.FontName = PropBag.ReadProperty("FontName", "굴림")
    txtBody.FontSize = PropBag.ReadProperty("FontSize", 10)
    txtBody.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    txtBody.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    m_FileName = PropBag.ReadProperty("FileName", m_def_FileName)
    txtBody.Text = PropBag.ReadProperty("Text", "")
    m_Dirty = PropBag.ReadProperty("Dirty", m_def_Dirty)
'    m_ErrNumber = PropBag.ReadProperty("ErrNumber", m_def_ErrNumber)
'    m_ErrDesc = PropBag.ReadProperty("ErrDesc", m_def_ErrDesc)
    m_ErrNumber = PropBag.ReadProperty("ErrNumber", m_def_ErrNumber)
    m_ErrDesc = PropBag.ReadProperty("ErrDesc", m_def_ErrDesc)
End Sub

'속성값을 저장소에 기록합니다.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("ForeColor", txtBody.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", txtBody.Enabled, True)
    Call PropBag.WriteProperty("FillColor", m_FillColor, m_def_FillColor)
    Call PropBag.WriteProperty("BackColor", txtBody.BackColor, &H80000005)
    Call PropBag.WriteProperty("FontBold", txtBody.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", txtBody.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", txtBody.FontName, "굴림")
    Call PropBag.WriteProperty("FontSize", txtBody.FontSize, 10)
    Call PropBag.WriteProperty("FontStrikethru", txtBody.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontUnderline", txtBody.FontUnderline, 0)
    Call PropBag.WriteProperty("FileName", m_FileName, m_def_FileName)
    Call PropBag.WriteProperty("Text", txtBody.Text, "")
    Call PropBag.WriteProperty("Dirty", m_Dirty, m_def_Dirty)
'    Call PropBag.WriteProperty("ErrNumber", m_ErrNumber, m_def_ErrNumber)
'    Call PropBag.WriteProperty("ErrDesc", m_ErrDesc, m_def_ErrDesc)
    Call PropBag.WriteProperty("ErrNumber", m_ErrNumber, m_def_ErrNumber)
    Call PropBag.WriteProperty("ErrDesc", m_ErrDesc, m_def_ErrDesc)
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=0,0,0,0
Public Property Get Dirty() As Boolean
Attribute Dirty.VB_Description = "열려진 파일에 변경 사항이 있는지 기록할 수 있습니다."
    Dirty = m_Dirty
End Property

Public Property Let Dirty(ByVal New_Dirty As Boolean)
    m_Dirty = New_Dirty
    PropertyChanged "Dirty"
End Property
'
''경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
''MemberInfo=7,0,0,0
'Public Property Get ErrNumber() As Integer
'    ErrNumber = m_ErrNumber
'End Property
'
'Public Property Let ErrNumber(ByVal New_ErrNumber As Integer)
'    m_ErrNumber = New_ErrNumber
'    PropertyChanged "ErrNumber"
'End Property
'
''경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
''MemberInfo=13,0,0,
'Public Property Get ErrDesc() As String
'    ErrDesc = m_ErrDesc
'End Property
'
'Public Property Let ErrDesc(ByVal New_ErrDesc As String)
'    m_ErrDesc = New_ErrDesc
'    PropertyChanged "ErrDesc"
'End Property
'
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=7,1,2,0
Public Property Get ErrNumber() As Integer
Attribute ErrNumber.VB_Description = "오류 발생시 오류 번호를 반환합니다."
Attribute ErrNumber.VB_MemberFlags = "400"
    ErrNumber = m_ErrNumber
End Property

Public Property Let ErrNumber(ByVal New_ErrNumber As Integer)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
Err.Raise 387
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,1,2,
Public Property Get ErrDesc() As String
Attribute ErrDesc.VB_Description = "오류 발생시 오류 설명을 반환합니다."
Attribute ErrDesc.VB_MemberFlags = "400"
    ErrDesc = m_ErrDesc
End Property

Public Property Let ErrDesc(ByVal New_ErrDesc As String)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
Err.Raise 387
End Property

