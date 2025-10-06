VERSION 5.00
Begin VB.Form frmAnalytics 
   Caption         =   "Analytics Dashboard"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   9840
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstBilling 
      Height          =   2175
      Left            =   6480
      TabIndex        =   3
      Top             =   360
      Width           =   2895
   End
   Begin VB.ListBox lstClaims 
      Height          =   2175
      Left            =   3480
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.ListBox lstPolicies 
      Height          =   2175
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.ListBox lstCompliance 
      Height          =   2655
      Left            =   480
      TabIndex        =   4
      Top             =   3120
      Width           =   2895
   End
   Begin VB.TextBox txtNarrative 
      Height          =   2655
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   3120
      Width           =   5895
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   360
      Left            =   8280
      TabIndex        =   0
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblBilling 
      Caption         =   "Billing Status"
      FontBold        =   -1  'True
      Height          =   255
      Left            =   6480
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblClaims 
      Caption         =   "Claim Status"
      FontBold        =   -1  'True
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblPolicies 
      Caption         =   "Policy Status"
      FontBold        =   -1  'True
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblCompliance 
      Caption         =   "Compliance Alerts"
      FontBold        =   -1  'True
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label lblNarrative 
      Caption         =   "Operational Summary"
      FontBold        =   -1  'True
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   2880
      Width           =   2175
   End
End
Attribute VB_Name = "frmAnalytics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRefresh_Click()
    LoadAnalytics
End Sub

Public Sub LoadAnalytics()
    LoadPolicyStatus
    LoadClaimStatus
    LoadBillingStatus
    LoadCompliance
    txtNarrative.Text = modReporting.BuildDashboardSummary
End Sub

Private Sub LoadPolicyStatus()
    Dim dict As Object
    Dim key As Variant
    Set dict = modReporting.GetPolicyStatusBreakdown
    lstPolicies.Clear
    For Each key In dict.Keys
        lstPolicies.AddItem key & ": " & dict(key)
    Next key
End Sub

Private Sub LoadClaimStatus()
    Dim dict As Object
    Dim key As Variant
    Set dict = modReporting.GetClaimStatusBreakdown
    lstClaims.Clear
    For Each key In dict.Keys
        lstClaims.AddItem key & ": " & dict(key)
    Next key
End Sub

Private Sub LoadBillingStatus()
    Dim dict As Object
    Dim key As Variant
    Set dict = modReporting.GetInvoiceStatusBreakdown
    lstBilling.Clear
    For Each key In dict.Keys
        lstBilling.AddItem key & ": " & dict(key)
    Next key
End Sub

Private Sub LoadCompliance()
    Dim alerts As Collection
    Dim item As Variant
    Set alerts = modReporting.GetComplianceAlerts
    lstCompliance.Clear
    For Each item In alerts
        lstCompliance.AddItem CStr(item)
    Next item
End Sub

Private Sub Form_Load()
    LoadAnalytics
End Sub
