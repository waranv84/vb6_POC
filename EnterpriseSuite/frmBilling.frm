VERSION 5.00
Begin VB.Form frmBilling 
   Caption         =   "Billing & Collections"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNarrative 
      Height          =   2175
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   4440
      Width           =   5655
   End
   Begin VB.CommandButton cmdMarkPaid 
      Caption         =   "Mark Paid"
      Height          =   360
      Left            =   7680
      TabIndex        =   8
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdApplyPayment 
      Caption         =   "Apply Payment"
      Height          =   360
      Left            =   5760
      TabIndex        =   7
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   360
      Left            =   3840
      TabIndex        =   6
      Top             =   3840
      Width           =   1575
   End
   Begin VB.ComboBox cmbStatus 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox txtPaymentEntry 
      Height          =   315
      Left            =   4560
      TabIndex        =   4
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox txtAmountPaid 
      Height          =   315
      Left            =   4560
      TabIndex        =   3
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtAmountDue 
      Height          =   315
      Left            =   4560
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtDueDate 
      Height          =   315
      Left            =   4560
      TabIndex        =   1
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txtInvoice 
      Height          =   315
      Left            =   4560
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin VB.ListBox lstInvoices 
      Height          =   5895
      Left            =   360
      TabIndex        =   10
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label lblNarrative 
      Caption         =   "Narrative"
      FontBold        =   -1  'True
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label lblPaymentEntry 
      Caption         =   "Payment Entry"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblAmountPaid 
      Caption         =   "Amount Paid"
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblAmountDue 
      Caption         =   "Amount Due"
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblDueDate 
      Caption         =   "Due Date"
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblInvoice 
      Caption         =   "Invoice #"
      Height          =   255
      Left            =   3600
      TabIndex        =   17
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mInvoices As Collection
Private mCurrent As clsInvoice

Private Sub cmdApplyPayment_Click()
    If mCurrent Is Nothing Then
        MsgBox "Select an invoice first.", vbInformation
        Exit Sub
    End If

    If Not IsNumeric(txtPaymentEntry.Text) Then
        MsgBox "Enter a numeric payment amount.", vbExclamation
        Exit Sub
    End If

    On Error GoTo HandleError
    modDatabase.RecordPayment mCurrent.InvoiceNumber, CCur(txtPaymentEntry.Text), cmbStatus.Text
    modUtilities.Log "Invoice " & mCurrent.InvoiceNumber & " payment applied"
    LoadInvoices
    SelectInvoice mCurrent.InvoiceNumber
    Exit Sub

HandleError:
    MsgBox "Unable to apply payment: " & Err.Description, vbCritical
End Sub

Private Sub cmdMarkPaid_Click()
    If mCurrent Is Nothing Then
        MsgBox "Select an invoice first.", vbInformation
        Exit Sub
    End If

    On Error GoTo HandleError
    modDatabase.RecordPayment mCurrent.InvoiceNumber, mCurrent.AmountDue, "Paid"
    modUtilities.Log "Invoice " & mCurrent.InvoiceNumber & " marked paid"
    LoadInvoices
    SelectInvoice mCurrent.InvoiceNumber
    Exit Sub

HandleError:
    MsgBox "Unable to mark paid: " & Err.Description, vbCritical
End Sub

Private Sub cmdRefresh_Click()
    LoadInvoices
End Sub

Private Sub Form_Load()
    PopulateStatuses
    LoadInvoices
End Sub

Private Sub lstInvoices_Click()
    If lstInvoices.ListIndex < 0 Then Exit Sub
    Set mCurrent = mInvoices(lstInvoices.ListIndex + 1)
    DisplayInvoice mCurrent
End Sub

Public Sub LoadInvoices()
    Set mInvoices = modDatabase.GetInvoices
    lstInvoices.Clear

    Dim idx As Integer
    For idx = 1 To mInvoices.Count
        Dim invoice As clsInvoice
        Set invoice = mInvoices(idx)
        lstInvoices.AddItem invoice.InvoiceNumber & " - " & invoice.Status
    Next idx

    If mInvoices.Count > 0 Then
        lstInvoices.ListIndex = 0
    Else
        ClearDetails
    End If
End Sub

Private Sub DisplayInvoice(ByVal invoice As clsInvoice)
    txtInvoice.Text = invoice.InvoiceNumber
    txtDueDate.Text = Format$(invoice.DueDate, "yyyy-mm-dd")
    txtAmountDue.Text = Format$(invoice.AmountDue, "0.00")
    txtAmountPaid.Text = Format$(invoice.AmountPaid, "0.00")
    txtPaymentEntry.Text = Format$(invoice.AmountDue - invoice.AmountPaid, "0.00")
    cmbStatus.Text = invoice.Status
    txtNarrative.Text = modReporting.BuildInvoiceNarrative(invoice)
    Set mCurrent = invoice
End Sub

Private Sub ClearDetails()
    txtInvoice.Text = ""
    txtDueDate.Text = ""
    txtAmountDue.Text = ""
    txtAmountPaid.Text = ""
    txtPaymentEntry.Text = ""
    cmbStatus.ListIndex = -1
    txtNarrative.Text = ""
    Set mCurrent = Nothing
End Sub

Private Sub PopulateStatuses()
    cmbStatus.Clear
    cmbStatus.AddItem "Paid"
    cmbStatus.AddItem "Due"
    cmbStatus.AddItem "Past Due"
    cmbStatus.AddItem "Payment Plan"
    cmbStatus.AddItem "Escalated"
    cmbStatus.ListIndex = 0
End Sub

Private Sub SelectInvoice(ByVal invoiceNumber As String)
    Dim idx As Integer
    For idx = 1 To mInvoices.Count
        If mInvoices(idx).InvoiceNumber = invoiceNumber Then
            lstInvoices.ListIndex = idx - 1
            Exit For
        End If
    Next idx
End Sub
