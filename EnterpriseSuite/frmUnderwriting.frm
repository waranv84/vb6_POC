VERSION 5.00
Begin VB.Form frmUnderwriting 
   Caption         =   "Underwriting Workbench"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   300
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNotes 
      Height          =   2055
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   4440
      Width           =   5655
   End
   Begin VB.CommandButton cmdEscalate 
      Caption         =   "Escalate"
      Height          =   360
      Left            =   7440
      TabIndex        =   7
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   360
      Left            =   5520
      TabIndex        =   6
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   360
      Left            =   3600
      TabIndex        =   5
      Top             =   3840
      Width           =   1575
   End
   Begin VB.ComboBox cmbStatus 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox txtRiskScore 
      Height          =   315
      Left            =   4560
      TabIndex        =   3
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox txtAssigned 
      Height          =   315
      Left            =   4560
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txtSubmission 
      Height          =   315
      Left            =   4560
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtPolicy 
      Height          =   315
      Left            =   4560
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
   Begin VB.ListBox lstQueue 
      Height          =   5895
      Left            =   360
      TabIndex        =   9
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label lblNotes 
      Caption         =   "Notes"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lblRiskScore 
      Caption         =   "Risk Score"
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblAssigned 
      Caption         =   "Assigned To"
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblSubmission 
      Caption         =   "Submission"
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblPolicy 
      Caption         =   "Policy"
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   1200
      Width           =   855
   End
End
Attribute VB_Name = "frmUnderwriting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCases As Collection
Private mCurrent As clsUnderwritingCase

Private Sub cmdEscalate_Click()
    If mCurrent Is Nothing Then
        MsgBox "Select a case first.", vbInformation
        Exit Sub
    End If

    If Not IsNumeric(txtRiskScore.Text) Then
        MsgBox "Enter a numeric risk score before escalating.", vbExclamation
        Exit Sub
    End If

    txtRiskScore.Text = CStr(CInt(txtRiskScore.Text) + 10)
    cmbStatus.Text = "Escalated"
    SaveCase
End Sub

Private Sub cmdRefresh_Click()
    LoadQueue
End Sub

Private Sub cmdUpdate_Click()
    SaveCase
End Sub

Private Sub Form_Load()
    PopulateStatuses
    LoadQueue
End Sub

Private Sub lstQueue_Click()
    If lstQueue.ListIndex < 0 Then Exit Sub
    Set mCurrent = mCases(lstQueue.ListIndex + 1)
    DisplayCase mCurrent
End Sub

Public Sub LoadQueue()
    Set mCases = modDatabase.GetUnderwritingCases
    lstQueue.Clear

    Dim idx As Integer
    For idx = 1 To mCases.Count
        Dim uw As clsUnderwritingCase
        Set uw = mCases(idx)
        lstQueue.AddItem uw.PolicyNumber & " - " & uw.Status
    Next idx

    If mCases.Count > 0 Then
        lstQueue.ListIndex = 0
    Else
        ClearDetails
    End If
End Sub

Private Sub DisplayCase(ByVal uw As clsUnderwritingCase)
    txtPolicy.Text = uw.PolicyNumber
    txtSubmission.Text = Format$(uw.SubmissionDate, "yyyy-mm-dd")
    txtAssigned.Text = uw.AssignedTo
    txtRiskScore.Text = CStr(uw.RiskScore)
    cmbStatus.Text = uw.Status
    txtNotes.Text = uw.Notes
    Set mCurrent = uw
End Sub

Private Sub ClearDetails()
    txtPolicy.Text = ""
    txtSubmission.Text = ""
    txtAssigned.Text = ""
    txtRiskScore.Text = ""
    cmbStatus.ListIndex = -1
    txtNotes.Text = ""
    Set mCurrent = Nothing
End Sub

Private Sub PopulateStatuses()
    cmbStatus.Clear
    cmbStatus.AddItem "Awaiting Documents"
    cmbStatus.AddItem "In Review"
    cmbStatus.AddItem "Escalated"
    cmbStatus.AddItem "Quoted"
    cmbStatus.AddItem "Declined"
    cmbStatus.AddItem "Bound"
    cmbStatus.ListIndex = 0
End Sub

Private Sub SaveCase()
    If mCurrent Is Nothing Then
        MsgBox "Select a case to update.", vbInformation
        Exit Sub
    End If

    If Not IsNumeric(txtRiskScore.Text) Then
        MsgBox "Enter a numeric risk score.", vbExclamation
        Exit Sub
    End If

    On Error GoTo HandleError
    modDatabase.UpdateUnderwritingCase txtPolicy.Text, cmbStatus.Text, CInt(txtRiskScore.Text)
    modUtilities.Log "Underwriting case " & txtPolicy.Text & " updated"
    LoadQueue
    SelectCase txtPolicy.Text
    Exit Sub

HandleError:
    MsgBox "Unable to update case: " & Err.Description, vbCritical
End Sub

Private Sub SelectCase(ByVal policyNumber As String)
    Dim idx As Integer
    For idx = 1 To mCases.Count
        If mCases(idx).PolicyNumber = policyNumber Then
            lstQueue.ListIndex = idx - 1
            Exit For
        End If
    Next idx
End Sub
