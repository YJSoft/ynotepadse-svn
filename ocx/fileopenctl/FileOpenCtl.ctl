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
      ScrollBars      =   3  '�����
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
'�⺻ �Ӽ� ��:
Const m_def_ErrNumber = 0
Const m_def_ErrDesc = "������ �����ϴ�."
'Const m_def_ErrNumber = 0
'Const m_def_ErrDesc = ""
Const m_def_Dirty = 0
Const m_def_FillColor = &H0&
Const m_def_FileName = "0"
'�Ӽ� ����:
Dim m_ErrNumber As Integer
Dim m_ErrDesc As String
'Dim m_ErrNumber As Integer
'Dim m_ErrDesc As String
Dim m_Dirty As Boolean
Dim m_FillColor As OLE_COLOR
Dim m_FileName As String
'�̺�Ʈ ����:
Event Click() 'MappingInfo=txtBody,txtBody,-1,Click
Attribute Click.VB_Description = "��ü���� ���콺 ���߸� �����ٰ� ���� �� �߻��մϴ�."
Event DblClick() 'MappingInfo=txtBody,txtBody,-1,DblClick
Attribute DblClick.VB_Description = "���콺 ���߸� ��ü���� ������ ���� �� �ٽ� ������ ������ �߻��մϴ�."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtBody,txtBody,-1,KeyDown
Attribute KeyDown.VB_Description = "��ü�� ��Ŀ���� ���� �� Ű�� ������ �߻��մϴ�."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtBody,txtBody,-1,KeyPress
Attribute KeyPress.VB_Description = "ANSIŰ�� ������ ������ ��� �߻��մϴ�."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtBody,txtBody,-1,KeyUp
Attribute KeyUp.VB_Description = "��ü�� ��Ŀ���� ���� �� Ű�� ������ �߻��մϴ�."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtBody,txtBody,-1,MouseDown
Attribute MouseDown.VB_Description = "��ü�� ��Ŀ���� ���� �� ���콺 ���߸� ������ �߻��մϴ�."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtBody,txtBody,-1,MouseMove
Attribute MouseMove.VB_Description = "���콺�� ������ ��� �߻��մϴ�."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtBody,txtBody,-1,MouseUp
Attribute MouseUp.VB_Description = "��ü�� ��Ŀ���� ���� �� ���콺 ���߸� ������ �߻��մϴ�."
Event Changed() 'MappingInfo=txtBody,txtBody,-1,Change
Attribute Changed.VB_Description = "��Ʈ���� ������ ����� �� �߻��մϴ�."



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
'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=txtBody,txtBody,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "��ü���� �ؽ�Ʈ�� �׷����� ǥ���ϴ� ������� ��ȯ�ϰų� �����մϴ�."
    ForeColor = txtBody.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtBody.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=txtBody,txtBody,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "����ڰ� ���� �̺�Ʈ�� ���� ��ü�� ������ �� �ִ����� ���θ� �����ϴ� ���� ��ȯ�ϰų� �����մϴ�."
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

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=10,3,2,&H0&
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "����, ��, ���ڸ� ä��� �� ���� ���� ��ȯ�ϰų� �����մϴ�."
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

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=txtBody,txtBody,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "��ü�� �ؽ�Ʈ�� �׷����� ǥ���ϱ� ���� ���Ǵ� ������ ��ȯ�ϰų� �����մϴ�."
    BackColor = txtBody.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtBody.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=txtBody,txtBody,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "���� �۲� ������ ��ȯ�ϰų� �����մϴ�."
Attribute FontBold.VB_ProcData.VB_Invoke_Property = "DefaultPage"
    FontBold = txtBody.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    txtBody.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=txtBody,txtBody,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "����� �۲� ������ ��ȯ�ϰų� �����մϴ�."
Attribute FontItalic.VB_ProcData.VB_Invoke_Property = "DefaultPage"
    FontItalic = txtBody.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    txtBody.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=txtBody,txtBody,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "�־��� �ܰ��� �� �࿡ ��Ÿ���� �۲��� �̸��� �����մϴ�."
Attribute FontName.VB_ProcData.VB_Invoke_Property = "DefaultPage"
    FontName = txtBody.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    txtBody.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=txtBody,txtBody,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "�־��� �ܰ��� �� �࿡ ��Ÿ���� �۲� ũ�⸦ ����Ʈ ������ �����մϴ�."
Attribute FontSize.VB_ProcData.VB_Invoke_Property = "DefaultPage"
    FontSize = txtBody.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    txtBody.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=txtBody,txtBody,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "��Ҽ� �۲� ������ ��ȯ�ϰų� �����մϴ�."
Attribute FontStrikethru.VB_ProcData.VB_Invoke_Property = "DefaultPage"
    FontStrikethru = txtBody.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    txtBody.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=txtBody,txtBody,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "���� �۲� ������ ��ȯ�ϰų� �����մϴ�."
Attribute FontUnderline.VB_ProcData.VB_Invoke_Property = "DefaultPage"
    FontUnderline = txtBody.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    txtBody.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=0
Public Function OpenFile(strFileName As String) As Boolean
Attribute OpenFile.VB_Description = "������ ���ϴ�.(���� �߻��� False���� ��ȯ�˴ϴ�)"
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

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=0
Public Function SaveFile(strFileName As String) As Boolean
Attribute SaveFile.VB_Description = "������ �����մϴ�.(���� �߻��� False���� ��ȯ�˴ϴ�)"
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

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,1,2,0
Public Property Get FileName() As String
Attribute FileName.VB_Description = "���� ������ ���� �̸��� ��ȯ�մϴ�."
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

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=txtBody,txtBody,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "��Ʈ�ѿ� ���Ե� �ؽ�Ʈ�� ��ȯ�ϰų� �����մϴ�."
    Text = txtBody.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtBody.Text() = New_Text
    PropertyChanged "Text"
End Property

'����� ���� ��Ʈ�ѿ� ���� �Ӽ��� �ʱ�ȭ�մϴ�.
Private Sub UserControl_InitProperties()
    m_FillColor = m_def_FillColor
    m_FileName = m_def_FileName
    m_Dirty = m_def_Dirty
'    m_ErrNumber = m_def_ErrNumber
'    m_ErrDesc = m_def_ErrDesc
    m_ErrNumber = m_def_ErrNumber
    m_ErrDesc = m_def_ErrDesc
End Sub

'����ҿ��� �Ӽ����� �ε��մϴ�.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    txtBody.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtBody.Enabled = PropBag.ReadProperty("Enabled", True)
    m_FillColor = PropBag.ReadProperty("FillColor", m_def_FillColor)
    txtBody.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtBody.FontBold = PropBag.ReadProperty("FontBold", 0)
    txtBody.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    txtBody.FontName = PropBag.ReadProperty("FontName", "����")
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

'�Ӽ����� ����ҿ� ����մϴ�.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("ForeColor", txtBody.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", txtBody.Enabled, True)
    Call PropBag.WriteProperty("FillColor", m_FillColor, m_def_FillColor)
    Call PropBag.WriteProperty("BackColor", txtBody.BackColor, &H80000005)
    Call PropBag.WriteProperty("FontBold", txtBody.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", txtBody.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", txtBody.FontName, "����")
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

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=0,0,0,0
Public Property Get Dirty() As Boolean
Attribute Dirty.VB_Description = "������ ���Ͽ� ���� ������ �ִ��� ����� �� �ֽ��ϴ�."
    Dirty = m_Dirty
End Property

Public Property Let Dirty(ByVal New_Dirty As Boolean)
    m_Dirty = New_Dirty
    PropertyChanged "Dirty"
End Property
'
''���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
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
''���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
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
'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=7,1,2,0
Public Property Get ErrNumber() As Integer
Attribute ErrNumber.VB_Description = "���� �߻��� ���� ��ȣ�� ��ȯ�մϴ�."
Attribute ErrNumber.VB_MemberFlags = "400"
    ErrNumber = m_ErrNumber
End Property

Public Property Let ErrNumber(ByVal New_ErrNumber As Integer)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
Err.Raise 387
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,1,2,
Public Property Get ErrDesc() As String
Attribute ErrDesc.VB_Description = "���� �߻��� ���� ������ ��ȯ�մϴ�."
Attribute ErrDesc.VB_MemberFlags = "400"
    ErrDesc = m_ErrDesc
End Property

Public Property Let ErrDesc(ByVal New_ErrDesc As String)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
Err.Raise 387
End Property

