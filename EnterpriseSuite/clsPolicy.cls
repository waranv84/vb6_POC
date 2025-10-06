VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPolicy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mPolicyNumber As String
Private mHolderName As String
Private mProductType As String
Private mEffectiveDate As Date
Private mExpirationDate As Date
Private mPremium As Currency
Private mStatus As String
Private mAgentCode As String

Public Property Get PolicyNumber() As String
    PolicyNumber = mPolicyNumber
End Property

Public Property Let PolicyNumber(ByVal value As String)
    mPolicyNumber = value
End Property

Public Property Get HolderName() As String
    HolderName = mHolderName
End Property

Public Property Let HolderName(ByVal value As String)
    mHolderName = value
End Property

Public Property Get ProductType() As String
    ProductType = mProductType
End Property

Public Property Let ProductType(ByVal value As String)
    mProductType = value
End Property

Public Property Get EffectiveDate() As Date
    EffectiveDate = mEffectiveDate
End Property

Public Property Let EffectiveDate(ByVal value As Date)
    mEffectiveDate = value
End Property

Public Property Get ExpirationDate() As Date
    ExpirationDate = mExpirationDate
End Property

Public Property Let ExpirationDate(ByVal value As Date)
    mExpirationDate = value
End Property

Public Property Get Premium() As Currency
    Premium = mPremium
End Property

Public Property Let Premium(ByVal value As Currency)
    mPremium = value
End Property

Public Property Get Status() As String
    Status = mStatus
End Property

Public Property Let Status(ByVal value As String)
    mStatus = value
End Property

Public Property Get AgentCode() As String
    AgentCode = mAgentCode
End Property

Public Property Let AgentCode(ByVal value As String)
    mAgentCode = value
End Property
