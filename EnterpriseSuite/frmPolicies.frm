VERSION 5.00
Begin VB.Form frmPolicies 
   Caption         =   "Policy Administration"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNarrative 
      Height          =   2055
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   4560
      Width           =   5535
   End
   Begin VB.CommandButton cmdUpdateStatus 
      Caption         =   "Update Status"
      Height          =   360
      Left            =   7560
      TabIndex        =   10
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save New Policy"
      Height          =   360
      Left            =   5760
      TabIndex        =   9
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   360
      Left            =   3960
      TabIndex        =   8
      Top             =   3960
      Width           =   1575
   End
   Begin VB.ComboBox cmbStatus 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox txtPremium 
      Height          =   315
      Left            =   4560
      TabIndex        =   6
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox txtExpiration 
      Height          =   315
      Left            =   4560
      TabIndex        =   5
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txtEffective 
      Height          =   315
      Left            =   4560
      TabIndex        =   4
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtProduct 
      Height          =   315
      Left            =   4560
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtHolder 
      Height          =   315
      Left            =   4560
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtPolicyNumber 
      Height          =   315
      Left            =   4560
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.ListBox lstPolicies 
      Height          =   5865
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label lblNarrativeTitle 
      Caption         =   "Narrative"
      FontBold        =   -1  'True
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lblPremium 
      Caption         =   "Premium"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lblExpiration 
      Caption         =   "Expiration"
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblEffective 
      Caption         =   "Effective"
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblProduct 
      Caption         =   "Product"
      Height          =   255
      Left            =   3600
      TabIndex        =   17
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblHolder 
      Caption         =   "Holder"
      Height          =   255
      Left            =   3600
      TabIndex        =   18
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblPolicyNumber 
      Caption         =   "Policy #"
      Height          =   255
      Left            =   3600
      TabIndex        =   19
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmPolicies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mPolicies As Collection
Private mCurrent As clsPolicy

Private Sub cmdRefresh_Click()
    LoadPolicies
End Sub

Private Sub cmdSave_Click()
    On Error GoTo HandleError

    Dim policy As clsPolicy
    Set policy = New clsPolicy
    policy.PolicyNumber = Trim$(txtPolicyNumber.Text)
    policy.HolderName = Trim$(txtHolder.Text)
    policy.ProductType = Trim$(txtProduct.Text)

    If policy.PolicyNumber = "" Or policy.HolderName = "" Then
        MsgBox "Policy number and holder are required.", vbExclamation
        Exit Sub
    End If

    If Not IsDate(txtEffective.Text) Or Not IsDate(txtExpiration.Text) Then
        MsgBox "Enter valid effective and expiration dates (e.g. 2024-06-30).", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(txtPremium.Text) Then
        MsgBox "Enter a numeric premium amount.", vbExclamation
        Exit Sub
    End If

    policy.EffectiveDate = CDate(txtEffective.Text)
    policy.ExpirationDate = CDate(txtExpiration.Text)
    policy.Premium = CCur(txtPremium.Text)
    policy.Status = cmbStatus.Text
    policy.AgentCode = "MANUAL"

    modDatabase.SavePolicy policy
    modUtilities.Log "Policy " & policy.PolicyNumber & " created"
    LoadPolicies
    SelectPolicyByNumber policy.PolicyNumber
    Exit Sub

HandleError:
    MsgBox "Unable to save policy: " & Err.Description, vbCritical
End Sub

Private Sub cmdUpdateStatus_Click()
    If mCurrent Is Nothing Then
        MsgBox "Select a policy to update.", vbInformation
        Exit Sub
    End If

    On Error GoTo HandleError
    If Not IsDate(txtExpiration.Text) Then
        MsgBox "Enter a valid expiration date before updating.", vbExclamation
        Exit Sub
    End If

    modDatabase.UpdatePolicyStatus mCurrent.PolicyNumber, cmbStatus.Text, CDate(txtExpiration.Text)
    modUtilities.Log "Policy " & mCurrent.PolicyNumber & " status updated"
    LoadPolicies
    SelectPolicyByNumber mCurrent.PolicyNumber
    Exit Sub

HandleError:
    MsgBox "Unable to update policy: " & Err.Description, vbCritical
End Sub

Private Sub Form_Load()
    PopulateStatuses
    LoadPolicies
End Sub

Private Sub lstPolicies_Click()
    If lstPolicies.ListIndex < 0 Then Exit Sub
    Set mCurrent = mPolicies(lstPolicies.ListIndex + 1)
    DisplayPolicy mCurrent
End Sub

Public Sub LoadPolicies()
    Set mPolicies = modDatabase.GetPolicies
    lstPolicies.Clear

    Dim idx As Integer
    For idx = 1 To mPolicies.Count
        Dim policy As clsPolicy
        Set policy = mPolicies(idx)
        lstPolicies.AddItem policy.PolicyNumber & " - " & policy.HolderName
    Next idx

    If mPolicies.Count > 0 Then
        lstPolicies.ListIndex = 0
    Else
        ClearDetails
    End If
End Sub

Private Sub DisplayPolicy(ByVal policy As clsPolicy)
    Set mCurrent = policy
    txtPolicyNumber.Text = policy.PolicyNumber
    txtHolder.Text = policy.HolderName
    txtProduct.Text = policy.ProductType
    txtEffective.Text = Format$(policy.EffectiveDate, "yyyy-mm-dd")
    txtExpiration.Text = Format$(policy.ExpirationDate, "yyyy-mm-dd")
    txtPremium.Text = Format$(policy.Premium, "0.00")
    cmbStatus.Text = policy.Status
    txtNarrative.Text = modReporting.BuildPolicyNarrative(policy)
End Sub

Private Sub ClearDetails()
    txtPolicyNumber.Text = ""
    txtHolder.Text = ""
    txtProduct.Text = ""
    txtEffective.Text = ""
    txtExpiration.Text = ""
    txtPremium.Text = ""
    cmbStatus.ListIndex = -1
    txtNarrative.Text = ""
    Set mCurrent = Nothing
End Sub

Private Sub PopulateStatuses()
    cmbStatus.Clear
    cmbStatus.AddItem "Active"
    cmbStatus.AddItem "Pending Renewal"
    cmbStatus.AddItem "Under Review"
    cmbStatus.AddItem "Cancelled"
    cmbStatus.AddItem "Non-Renewing"
    cmbStatus.AddItem "Prospect"
    cmbStatus.ListIndex = 0
End Sub

Private Sub SelectPolicyByNumber(ByVal policyNumber As String)
    Dim idx As Integer
    For idx = 1 To mPolicies.Count
        If mPolicies(idx).PolicyNumber = policyNumber Then
            lstPolicies.ListIndex = idx - 1
            Exit For
        End If
    Next idx
End Sub
