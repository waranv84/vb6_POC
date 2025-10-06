VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUnderwritingCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mPolicyNumber As String
Private mSubmissionDate As Date
Private mRiskScore As Integer
Private mAssignedTo As String
Private mStatus As String
Private mNotes As String

Public Property Get PolicyNumber() As String
    PolicyNumber = mPolicyNumber
End Property

Public Property Let PolicyNumber(ByVal value As String)
    mPolicyNumber = value
End Property

Public Property Get SubmissionDate() As Date
    SubmissionDate = mSubmissionDate
End Property

Public Property Let SubmissionDate(ByVal value As Date)
    mSubmissionDate = value
End Property

Public Property Get RiskScore() As Integer
    RiskScore = mRiskScore
End Property

Public Property Let RiskScore(ByVal value As Integer)
    mRiskScore = value
End Property

Public Property Get AssignedTo() As String
    AssignedTo = mAssignedTo
End Property

Public Property Let AssignedTo(ByVal value As String)
    mAssignedTo = value
End Property

Public Property Get Status() As String
    Status = mStatus
End Property

Public Property Let Status(ByVal value As String)
    mStatus = value
End Property

Public Property Get Notes() As String
    Notes = mNotes
End Property

Public Property Let Notes(ByVal value As String)
    mNotes = value
End Property
