VERSION 5.00
Begin VB.Form frmDocuments 
   Caption         =   "Document Compliance"
   ClientHeight    =   6690
   ClientLeft      =   90
   ClientTop       =   360
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSummary 
      Height          =   2055
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4200
      Width           =   5655
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   360
      Left            =   5760
      TabIndex        =   6
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   360
      Left            =   3840
      TabIndex        =   5
      Top             =   3600
      Width           =   1575
   End
   Begin VB.ComboBox cmbCompliance 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox txtReceived 
      Height          =   315
      Left            =   4560
      TabIndex        =   3
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox txtType 
      Height          =   315
      Left            =   4560
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   4560
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtPolicy 
      Height          =   315
      Left            =   4560
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
   Begin VB.ListBox lstDocuments 
      Height          =   5895
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label lblSummary 
      Caption         =   "Summary"
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lblCompliance 
      Caption         =   "Compliance"
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lblReceived 
      Caption         =   "Received"
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblType 
      Caption         =   "Type"
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblName 
      Caption         =   "Document"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblPolicy 
      Caption         =   "Policy"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "frmDocuments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mDocuments As Collection
Private mCurrent As clsDocument

Private Sub cmdRefresh_Click()
    LoadDocuments
End Sub

Private Sub cmdUpdate_Click()
    If mCurrent Is Nothing Then
        MsgBox "Select a document first.", vbInformation
        Exit Sub
    End If

    On Error GoTo HandleError
    modDatabase.UpdateDocumentCompliance txtPolicy.Text, txtName.Text, cmbCompliance.Text
    modUtilities.Log "Document " & txtName.Text & " marked " & cmbCompliance.Text
    LoadDocuments
    SelectDocument txtPolicy.Text, txtName.Text
    Exit Sub

HandleError:
    MsgBox "Unable to update document: " & Err.Description, vbCritical
End Sub

Private Sub Form_Load()
    PopulateStatuses
    LoadDocuments
End Sub

Private Sub lstDocuments_Click()
    If lstDocuments.ListIndex < 0 Then Exit Sub
    Set mCurrent = mDocuments(lstDocuments.ListIndex + 1)
    DisplayDocument mCurrent
End Sub

Public Sub LoadDocuments()
    Set mDocuments = modDatabase.GetDocuments
    lstDocuments.Clear

    Dim idx As Integer
    For idx = 1 To mDocuments.Count
        Dim doc As clsDocument
        Set doc = mDocuments(idx)
        lstDocuments.AddItem doc.PolicyNumber & " - " & doc.DocumentName & " (" & doc.ComplianceStatus & ")"
    Next idx

    If mDocuments.Count > 0 Then
        lstDocuments.ListIndex = 0
    Else
        ClearDetails
    End If
End Sub

Private Sub DisplayDocument(ByVal doc As clsDocument)
    txtPolicy.Text = doc.PolicyNumber
    txtName.Text = doc.DocumentName
    txtType.Text = doc.DocumentType
    txtReceived.Text = Format$(doc.ReceivedDate, "yyyy-mm-dd")
    cmbCompliance.Text = doc.ComplianceStatus
    txtSummary.Text = "Policy " & doc.PolicyNumber & " submitted " & doc.DocumentName & " (" & doc.DocumentType & ")" & vbCrLf & _
                     "Received " & Format$(doc.ReceivedDate, "yyyy-mm-dd") & " | Compliance: " & doc.ComplianceStatus
    Set mCurrent = doc
End Sub

Private Sub ClearDetails()
    txtPolicy.Text = ""
    txtName.Text = ""
    txtType.Text = ""
    txtReceived.Text = ""
    cmbCompliance.ListIndex = -1
    txtSummary.Text = ""
    Set mCurrent = Nothing
End Sub

Private Sub PopulateStatuses()
    cmbCompliance.Clear
    cmbCompliance.AddItem "Approved"
    cmbCompliance.AddItem "Under Review"
    cmbCompliance.AddItem "Rejected"
    cmbCompliance.AddItem "Expired"
    cmbCompliance.AddItem "Missing"
    cmbCompliance.ListIndex = 0
End Sub

Private Sub SelectDocument(ByVal policyNumber As String, ByVal docName As String)
    Dim idx As Integer
    For idx = 1 To mDocuments.Count
        If mDocuments(idx).PolicyNumber = policyNumber And mDocuments(idx).DocumentName = docName Then
            lstDocuments.ListIndex = idx - 1
            Exit For
        End If
    Next idx
End Sub
