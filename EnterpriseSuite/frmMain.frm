VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Insurance Operations Suite"
   ClientHeight    =   7080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh Dashboard"
      Height          =   360
      Left            =   7440
      TabIndex        =   3
      Top             =   360
      Width           =   2235
   End
   Begin VB.ListBox lstAudit 
      Height          =   2235
      Left            =   5160
      TabIndex        =   2
      Top             =   4560
      Width           =   4455
   End
   Begin VB.ListBox lstCompliance 
      Height          =   2235
      Left            =   360
      TabIndex        =   1
      Top             =   4560
      Width           =   4455
   End
   Begin VB.PictureBox picDashboard 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   360
      ScaleHeight     =   3465
      ScaleWidth      =   9255
      TabIndex        =   0
      Top             =   840
      Width           =   9285
      Begin VB.Label lblSummary 
         Caption         =   "Dashboard metrics will appear here."
         Height          =   2970
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   8775
      End
      Begin VB.Label lblDashboardTitle 
         Caption         =   "Operational Snapshot"
         FontBold        =   -1  'True
         FontSize        =   16
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   0
         Width           =   3495
      End
   End
   Begin VB.Label lblAuditTitle 
      Caption         =   "Recent Audit Activity"
      FontBold        =   -1  'True
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label lblComplianceTitle 
      Caption         =   "Compliance Alerts"
      FontBold        =   -1  'True
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuModules 
      Caption         =   "&Modules"
      Begin VB.Menu mnuPolicies 
         Caption         =   "&Policies"
      End
      Begin VB.Menu mnuClaims 
         Caption         =   "&Claims"
      End
      Begin VB.Menu mnuBilling 
         Caption         =   "&Billing"
      End
      Begin VB.Menu mnuAgents 
         Caption         =   "&Agents"
      End
      Begin VB.Menu mnuUnderwriting 
         Caption         =   "&Underwriting"
      End
      Begin VB.Menu mnuDocuments 
         Caption         =   "&Documents"
      End
      Begin VB.Menu mnuAnalytics 
         Caption         =   "&Analytics"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRefresh_Click()
    RefreshDashboard
End Sub

Private Sub Form_Load()
    modDatabase.InitializeDatabase
    RefreshDashboard
End Sub

Private Sub RefreshDashboard()
    lstCompliance.Clear
    lstAudit.Clear
    lblSummary.Caption = modReporting.BuildDashboardSummary

    Dim alerts As Collection
    Dim item As Variant
    Set alerts = modReporting.GetComplianceAlerts
    For Each item In alerts
        lstCompliance.AddItem CStr(item)
    Next item

    Dim audit As Collection
    Set audit = modReporting.GetRecentAuditTrail(8)
    For Each item In audit
        lstAudit.AddItem CStr(item)
    Next item

    modUtilities.Log "Dashboard refreshed"
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuAgents_Click()
    frmAgents.Show
    frmAgents.LoadAgents
End Sub

Private Sub mnuAnalytics_Click()
    frmAnalytics.Show
    frmAnalytics.LoadAnalytics
End Sub

Private Sub mnuBilling_Click()
    frmBilling.Show
    frmBilling.LoadInvoices
End Sub

Private Sub mnuClaims_Click()
    frmClaims.Show
    frmClaims.LoadClaims
End Sub

Private Sub mnuDocuments_Click()
    frmDocuments.Show
    frmDocuments.LoadDocuments
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuPolicies_Click()
    frmPolicies.Show
    frmPolicies.LoadPolicies
End Sub

Private Sub mnuSettings_Click()
    frmSettings.Show vbModal
    RefreshDashboard
End Sub

Private Sub mnuUnderwriting_Click()
    frmUnderwriting.Show
    frmUnderwriting.LoadQueue
End Sub
