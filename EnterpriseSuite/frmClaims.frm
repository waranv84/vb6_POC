VERSION 5.00
Begin VB.Form frmClaims 
   Caption         =   "Claim Management"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNarrative 
      Height          =   2175
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   4320
      Width           =   5655
   End
   Begin VB.CommandButton cmdAdjust 
      Caption         =   "Adjust Reserve"
      Height          =   360
      Left            =   7680
      TabIndex        =   8
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdvance 
      Caption         =   "Advance Status"
      Height          =   360
      Left            =   5760
      TabIndex        =   7
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   360
      Left            =   3840
      TabIndex        =   6
      Top             =   3720
      Width           =   1575
   End
   Begin VB.ComboBox cmbStatus 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox txtAmount 
      Height          =   315
      Left            =   4560
      TabIndex        =   4
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox txtReported 
      Height          =   315
      Left            =   4560
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txtLossDate 
      Height          =   315
      Left            =   4560
      TabIndex        =   2
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtPolicy 
      Height          =   315
      Left            =   4560
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtClaim 
      Height          =   315
      Left            =   4560
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.ListBox lstClaims 
      Height          =   5895
      Left            =   360
      TabIndex        =   10
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label lblNarrative 
      Caption         =   "Narrative"
      FontBold        =   -1  'True
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lblAmount 
      Caption         =   "Amount"
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lblReported 
      Caption         =   "Reported"
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblLossDate 
      Caption         =   "Loss Date"
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblPolicy 
      Caption         =   "Policy"
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblClaim 
      Caption         =   "Claim #"
      Height          =   255
      Left            =   3600
      TabIndex        =   17
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "frmClaims"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mClaims As Collection
Private mCurrent As clsClaim

Private Sub cmdAdjust_Click()
    UpdateClaim
End Sub

Private Sub cmdAdvance_Click()
    UpdateClaim
End Sub

Private Sub cmdRefresh_Click()
    LoadClaims
End Sub

Private Sub Form_Load()
    PopulateStatuses
    LoadClaims
End Sub

Private Sub lstClaims_Click()
    If lstClaims.ListIndex < 0 Then Exit Sub
    Set mCurrent = mClaims(lstClaims.ListIndex + 1)
    DisplayClaim mCurrent
End Sub

Public Sub LoadClaims()
    Set mClaims = modDatabase.GetClaims
    lstClaims.Clear

    Dim idx As Integer
    For idx = 1 To mClaims.Count
        Dim claim As clsClaim
        Set claim = mClaims(idx)
        lstClaims.AddItem claim.ClaimNumber & " - " & claim.Status
    Next idx

    If mClaims.Count > 0 Then
        lstClaims.ListIndex = 0
    Else
        ClearDetails
    End If
End Sub

Private Sub DisplayClaim(ByVal claim As clsClaim)
    txtClaim.Text = claim.ClaimNumber
    txtPolicy.Text = claim.PolicyNumber
    txtLossDate.Text = Format$(claim.LossDate, "yyyy-mm-dd")
    txtReported.Text = Format$(claim.ReportedDate, "yyyy-mm-dd")
    txtAmount.Text = Format$(claim.Amount, "0.00")
    cmbStatus.Text = claim.Status
    txtNarrative.Text = modReporting.BuildClaimNarrative(claim)
    Set mCurrent = claim
End Sub

Private Sub ClearDetails()
    txtClaim.Text = ""
    txtPolicy.Text = ""
    txtLossDate.Text = ""
    txtReported.Text = ""
    txtAmount.Text = ""
    cmbStatus.ListIndex = -1
    txtNarrative.Text = ""
    Set mCurrent = Nothing
End Sub

Private Sub PopulateStatuses()
    cmbStatus.Clear
    cmbStatus.AddItem "Submitted"
    cmbStatus.AddItem "Investigation"
    cmbStatus.AddItem "Awaiting Settlement"
    cmbStatus.AddItem "Approved"
    cmbStatus.AddItem "Closed"
    cmbStatus.AddItem "Denied"
    cmbStatus.ListIndex = 0
End Sub

Private Sub UpdateClaim()
    If mCurrent Is Nothing Then
        MsgBox "Select a claim to update.", vbInformation
        Exit Sub
    End If

    If Not IsNumeric(txtAmount.Text) Then
        MsgBox "Enter a numeric reserve amount.", vbExclamation
        Exit Sub
    End If

    On Error GoTo HandleError
    modDatabase.RecordClaimProgress mCurrent.ClaimNumber, cmbStatus.Text, CCur(txtAmount.Text)
    modUtilities.Log "Claim " & mCurrent.ClaimNumber & " updated"
    LoadClaims
    SelectClaimByNumber mCurrent.ClaimNumber
    Exit Sub

HandleError:
    MsgBox "Unable to update claim: " & Err.Description, vbCritical
End Sub

Private Sub SelectClaimByNumber(ByVal claimNumber As String)
    Dim idx As Integer
    For idx = 1 To mClaims.Count
        If mClaims(idx).ClaimNumber = claimNumber Then
            lstClaims.ListIndex = idx - 1
            Exit For
        End If
    Next idx
End Sub
