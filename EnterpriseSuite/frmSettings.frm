VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Suite Settings"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   360
      Left            =   4320
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdPurgeLog 
      Caption         =   "Clear Activity Log"
      Height          =   360
      Left            =   2160
      TabIndex        =   2
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton cmdOpenData 
      Caption         =   "Open Data Folder"
      Height          =   360
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtDbPath 
      Height          =   345
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   5160
   End
   Begin VB.Label lblDbPath 
      Caption         =   "Access database location"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label lblInfo 
      Caption         =   "Manage application storage and logging."
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOpenData_Click()
    MsgBox "Open this folder in Windows Explorer: " & modUtilities.AppDataPath, vbInformation
End Sub

Private Sub cmdPurgeLog_Click()
    On Error Resume Next
    Open modUtilities.CombinePath(App.Path, "InsuranceSuite.log") For Output As #1
    Close #1
    MsgBox "Activity log cleared.", vbInformation
End Sub

Private Sub Form_Load()
    txtDbPath.Text = modUtilities.CombinePath(modUtilities.AppDataPath, "InsuranceSuite.mdb")
End Sub
