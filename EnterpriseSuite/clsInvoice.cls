VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mInvoiceNumber As String
Private mPolicyNumber As String
Private mDueDate As Date
Private mAmountDue As Currency
Private mAmountPaid As Currency
Private mStatus As String
Private mLastUpdated As Date

Public Property Get InvoiceNumber() As String
    InvoiceNumber = mInvoiceNumber
End Property

Public Property Let InvoiceNumber(ByVal value As String)
    mInvoiceNumber = value
End Property

Public Property Get PolicyNumber() As String
    PolicyNumber = mPolicyNumber
End Property

Public Property Let PolicyNumber(ByVal value As String)
    mPolicyNumber = value
End Property

Public Property Get DueDate() As Date
    DueDate = mDueDate
End Property

Public Property Let DueDate(ByVal value As Date)
    mDueDate = value
End Property

Public Property Get AmountDue() As Currency
    AmountDue = mAmountDue
End Property

Public Property Let AmountDue(ByVal value As Currency)
    mAmountDue = value
End Property

Public Property Get AmountPaid() As Currency
    AmountPaid = mAmountPaid
End Property

Public Property Let AmountPaid(ByVal value As Currency)
    mAmountPaid = value
End Property

Public Property Get Status() As String
    Status = mStatus
End Property

Public Property Let Status(ByVal value As String)
    mStatus = value
End Property

Public Property Get LastUpdated() As Date
    LastUpdated = mLastUpdated
End Property

Public Property Let LastUpdated(ByVal value As Date)
    mLastUpdated = value
End Property
