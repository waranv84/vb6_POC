VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Insurance Operations Suite"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   360
      Left            =   4200
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Caption         =   "Insurance Operations Suite demonstrates a multi-department VB6 desktop application for managing policies, claims, billing, compliance, and underwriting using a self-contained Access database."
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5160
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub
