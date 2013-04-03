VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Option Base 0

Implements IBindStatusCallback

Private Sub IBindStatusCallback_GetBindInfo(grfBINDF As olelib.BINDF, pbindinfo As olelib.BINDINFO)

End Sub

Private Function IBindStatusCallback_GetPriority() As Long

End Function

Private Sub IBindStatusCallback_OnDataAvailable(ByVal grfBSCF As olelib.BSCF, ByVal dwSize As Long, pformatetc As olelib.FORMATETC, pStgmed As olelib.STGMEDIUM)

End Sub

Private Sub IBindStatusCallback_OnLowResource(ByVal reserved As Long)

End Sub

Private Sub IBindStatusCallback_OnObjectAvailable(riid As olelib.UUID, ByVal pUnk As stdole.IUnknown)

End Sub

Private Sub IBindStatusCallback_OnProgress(ByVal ulProgress As Long, ByVal ulProgressMax As Long, ByVal ulStatusCode As olelib.BINDSTATUS, ByVal szStatusText As Long)

End Sub

Private Sub IBindStatusCallback_OnStartBinding(ByVal dwReserved As Long, ByVal pib As olelib.IBinding)

End Sub

Private Sub IBindStatusCallback_OnStopBinding(ByVal hresult As Long, ByVal szError As Long)

End Sub
