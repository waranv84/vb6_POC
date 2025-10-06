VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mPolicyNumber As String
Private mDocumentName As String
Private mDocumentType As String
Private mReceivedDate As Date
Private mComplianceStatus As String

Public Property Get PolicyNumber() As String
    PolicyNumber = mPolicyNumber
End Property

Public Property Let PolicyNumber(ByVal value As String)
    mPolicyNumber = value
End Property

Public Property Get DocumentName() As String
    DocumentName = mDocumentName
End Property

Public Property Let DocumentName(ByVal value As String)
    mDocumentName = value
End Property

Public Property Get DocumentType() As String
    DocumentType = mDocumentType
End Property

Public Property Let DocumentType(ByVal value As String)
    mDocumentType = value
End Property

Public Property Get ReceivedDate() As Date
    ReceivedDate = mReceivedDate
End Property

Public Property Let ReceivedDate(ByVal value As Date)
    mReceivedDate = value
End Property

Public Property Get ComplianceStatus() As String
    ComplianceStatus = mComplianceStatus
End Property

Public Property Let ComplianceStatus(ByVal value As String)
    mComplianceStatus = value
End Property
