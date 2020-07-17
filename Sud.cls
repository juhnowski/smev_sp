VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pName As String
Private psudyaName As Collection
Private pAddress As String
Private pSite As String

Property Get site() As String
    site = pSite
End Property

Property Let site(s As String)
    pSite = s
End Property


Property Get name() As String
    name = pName
End Property

Property Let name(s As String)
    pName = s
End Property

Property Get sudyaName() As Collection
    Set sudyaName = psudyaName
End Property

Property Let sudyaName(s As Collection)
    Set psudyaName = s
End Property

Property Get address() As String
    address = pAddress
End Property

Property Let address(s As String)
    pAddress = s
End Property

Private Sub Class_Initialize()
    Set psudyaName = New Collection
End Sub

