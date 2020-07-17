VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Vziskaniya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pVidRashodov As String
Private pSumma As String

Property Get VidRashodov() As String
    VidRashodov = pVidRashodov
End Property

Property Let VidRashodov(s As String)
    pVidRashodov = s
End Property

Property Get summa() As String
    summa = pSumma
End Property

Property Let summa(s As String)
    pSumma = s
End Property
