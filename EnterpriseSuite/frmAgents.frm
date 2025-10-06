VERSION 5.00
Begin VB.Form frmAgents 
   Caption         =   "Agency Management"
   ClientHeight    =   6750
   ClientLeft      =   90
   ClientTop       =   360
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Agent"
      Height          =   360
      Left            =   7320
      TabIndex        =   9
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdExtend 
      Caption         =   "Extend License"
      Height          =   360
      Left            =   5400
      TabIndex        =   8
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   360
      Left            =   3480
      TabIndex        =   7
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CheckBox chkActive 
      Caption         =   "Active"
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtLicense 
      Height          =   315
      Left            =   4560
      TabIndex        =   5
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox txtPhone 
      Height          =   315
      Left            =   4560
      TabIndex        =   4
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtEmail 
      Height          =   315
      Left            =   4560
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtRegion 
      Height          =   315
      Left            =   4560
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   4560
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txtCode 
      Height          =   315
      Left            =   4560
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.ListBox lstAgents 
      Height          =   5865
      Left            =   360
      TabIndex        =   10
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblLicense 
      Caption         =   "License Expiration"
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblPhone 
      Caption         =   "Phone"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblEmail 
      Caption         =   "Email"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblRegion 
      Caption         =   "Region"
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblCode 
      Caption         =   "Agent Code"
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmAgents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mAgents As Collection
Private mCurrent As clsAgent

Private Sub chkActive_Click()
    If mCurrent Is Nothing Then Exit Sub
    mCurrent.Active = (chkActive.Value = vbChecked)
End Sub

Private Sub cmdExtend_Click()
    If mCurrent Is Nothing Then
        MsgBox "Select an agent first.", vbInformation
        Exit Sub
    End If

    If Not IsDate(txtLicense.Text) Then
        MsgBox "Enter a valid license expiration date before extending.", vbExclamation
        Exit Sub
    End If

    txtLicense.Text = Format$(DateAdd("m", 12, CDate(txtLicense.Text)), "yyyy-mm-dd")
    SaveAgent
End Sub

Private Sub cmdRefresh_Click()
    LoadAgents
End Sub

Private Sub cmdSave_Click()
    SaveAgent
End Sub

Private Sub Form_Load()
    LoadAgents
End Sub

Private Sub lstAgents_Click()
    If lstAgents.ListIndex < 0 Then Exit Sub
    Set mCurrent = mAgents(lstAgents.ListIndex + 1)
    DisplayAgent mCurrent
End Sub

Public Sub LoadAgents()
    Set mAgents = modDatabase.GetAgents
    lstAgents.Clear

    Dim idx As Integer
    For idx = 1 To mAgents.Count
        Dim agent As clsAgent
        Set agent = mAgents(idx)
        Dim statusLabel As String
        If agent.Active Then
            statusLabel = "Active"
        Else
            statusLabel = "Inactive"
        End If
        lstAgents.AddItem agent.AgentCode & " - " & agent.FullName & " (" & statusLabel & ")"
    Next idx

    If mAgents.Count > 0 Then
        lstAgents.ListIndex = 0
    Else
        ClearDetails
    End If
End Sub

Private Sub DisplayAgent(ByVal agent As clsAgent)
    txtCode.Text = agent.AgentCode
    txtName.Text = agent.FullName
    txtRegion.Text = agent.Region
    txtEmail.Text = agent.Email
    txtPhone.Text = agent.Phone
    txtLicense.Text = Format$(agent.LicenseExpiration, "yyyy-mm-dd")
    chkActive.Value = IIf(agent.Active, vbChecked, vbUnchecked)
    Set mCurrent = agent
End Sub

Private Sub ClearDetails()
    txtCode.Text = ""
    txtName.Text = ""
    txtRegion.Text = ""
    txtEmail.Text = ""
    txtPhone.Text = ""
    txtLicense.Text = ""
    chkActive.Value = vbUnchecked
    Set mCurrent = Nothing
End Sub

Private Sub SaveAgent()
    If mCurrent Is Nothing Then
        MsgBox "Select an agent to update.", vbInformation
        Exit Sub
    End If

    If Not IsDate(txtLicense.Text) Then
        MsgBox "Enter a valid license expiration date (e.g. 2024-12-31).", vbExclamation
        Exit Sub
    End If

    On Error GoTo HandleError
    modDatabase.UpdateAgent txtCode.Text, (chkActive.Value = vbChecked), CDate(txtLicense.Text)
    modUtilities.Log "Agent " & txtCode.Text & " updated"
    LoadAgents
    SelectAgent txtCode.Text
    Exit Sub

HandleError:
    MsgBox "Unable to update agent: " & Err.Description, vbCritical
End Sub

Private Sub SelectAgent(ByVal agentCode As String)
    Dim idx As Integer
    For idx = 1 To mAgents.Count
        If mAgents(idx).AgentCode = agentCode Then
            lstAgents.ListIndex = idx - 1
            Exit For
        End If
    Next idx
End Sub
