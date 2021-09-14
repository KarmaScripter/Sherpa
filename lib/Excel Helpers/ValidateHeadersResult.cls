VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidateHeadersResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private myResult As Boolean
Private implHeaders As Collection
Private missedHeaders As Collection
Private addtlHeaders As Collection
Private Header As Variant

Private Sub Class_Initialize()
    myResult = True
    Set implHeaders = New Collection
    Set missedHeaders = New Collection
    Set addtlHeaders = New Collection
End Sub

Public Property Get HasRequiredHeaders() As Boolean
    HasRequiredHeaders = myResult
End Property

Public Property Let HasRequiredHeaders(ByVal Value As Boolean)
    myResult = Value
End Property

Public Sub hasHeader(ByVal Value As String)
    implHeaders.Add Value
End Sub

Public Property Get ImplementedHeaders() As Collection
    Dim returnVal As New Collection
    For Each Header In implHeaders
        returnVal.Add Header
    Next Header
    Set ImplementedHeaders = returnVal
End Property

Public Sub HeaderNotFound(ByVal Value As String)
    missedHeaders.Add Value
End Sub

Public Property Get MissingHeaders() As Collection
    Dim returnVal As New Collection
    For Each Header In missedHeaders
        returnVal.Add Header
    Next Header
    Set MissingHeaders = returnVal
End Property

Public Sub HeaderExtra(ByVal Value As String)
    addtlHeaders.Add Value
End Sub

Public Property Get ExtraHeaders() As Collection
    Dim returnVal As New Collection
    For Each Header In addtlHeaders
        returnVal.Add Header
    Next Header
    Set ExtraHeaders = returnVal
End Property
