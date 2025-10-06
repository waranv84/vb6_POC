VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mAgentCode As String
Private mFullName As String
Private mRegion As String
Private mEmail As String
Private mPhone As String
Private mLicenseExpiration As Date
Private mActive As Boolean

Public Property Get AgentCode() As String
    AgentCode = mAgentCode
End Property

Public Property Let AgentCode(ByVal value As String)
    mAgentCode = value
End Property

Public Property Get FullName() As String
    FullName = mFullName
End Property

Public Property Let FullName(ByVal value As String)
    mFullName = value
End Property

Public Property Get Region() As String
    Region = mRegion
End Property

Public Property Let Region(ByVal value As String)
    mRegion = value
End Property

Public Property Get Email() As String
    Email = mEmail
End Property

Public Property Let Email(ByVal value As String)
    mEmail = value
End Property

Public Property Get Phone() As String
    Phone = mPhone
End Property

Public Property Let Phone(ByVal value As String)
    mPhone = value
End Property

Public Property Get LicenseExpiration() As Date
    LicenseExpiration = mLicenseExpiration
End Property

Public Property Let LicenseExpiration(ByVal value As Date)
    mLicenseExpiration = value
End Property

Public Property Get Active() As Boolean
    Active = mActive
End Property

Public Property Let Active(ByVal value As Boolean)
    mActive = value
End Property
