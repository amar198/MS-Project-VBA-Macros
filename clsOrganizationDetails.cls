VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOrganizationDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim sOrganisationName As String
Dim sManagerName As String

Public Property Get OrganisationName() As String
    OrganisationName = sOrganisationName
End Property

Private Sub Class_Initialize()
    sOrganisationName = "MyOrganisation Name"
End Sub

Public Property Get GetProjectManagerName() As String
    GetProjectManagerName = sManagerName
End Property

Public Property Let GetProjectManagerName(ByVal sManagerNm As String)
    sManagerName = sManagerNm
End Property
