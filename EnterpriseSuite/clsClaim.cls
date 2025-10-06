VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mClaimNumber As String
Private mPolicyNumber As String
Private mLossDate As Date
Private mReportedDate As Date
Private mAmount As Currency
Private mStatus As String
Private mAdjusterName As String
Private mDescription As String

Public Property Get ClaimNumber() As String
    ClaimNumber = mClaimNumber
End Property

Public Property Let ClaimNumber(ByVal value As String)
    mClaimNumber = value
End Property

Public Property Get PolicyNumber() As String
    PolicyNumber = mPolicyNumber
End Property

Public Property Let PolicyNumber(ByVal value As String)
    mPolicyNumber = value
End Property

Public Property Get LossDate() As Date
    LossDate = mLossDate
End Property

Public Property Let LossDate(ByVal value As Date)
    mLossDate = value
End Property

Public Property Get ReportedDate() As Date
    ReportedDate = mReportedDate
End Property

Public Property Let ReportedDate(ByVal value As Date)
    mReportedDate = value
End Property

Public Property Get Amount() As Currency
    Amount = mAmount
End Property

Public Property Let Amount(ByVal value As Currency)
    mAmount = value
End Property

Public Property Get Status() As String
    Status = mStatus
End Property

Public Property Let Status(ByVal value As String)
    mStatus = value
End Property

Public Property Get AdjusterName() As String
    AdjusterName = mAdjusterName
End Property

Public Property Let AdjusterName(ByVal value As String)
    mAdjusterName = value
End Property

Public Property Get Description() As String
    Description = mDescription
End Property

Public Property Let Description(ByVal value As String)
    mDescription = value
End Property
