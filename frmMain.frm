VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Signature Updater"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4830
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraContent 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      Begin VB.Label lblContent 
         Height          =   2295
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update Now"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdUpdate_Click()
    Call UpdateSignature
End Sub

Private Sub Form_Load()
    Call LoadConfigFromFile(Command)
    Me.Caption = APP_NAME & " - " & app_configuration
    
    If (config_autoclose = 1) Then
        Call UpdateSignature
        Unload Me
    End If
End Sub

Public Sub UpdateSignature()
    Me.lblContent.Caption = RssLastEntry(HttpGetRequest(config_feedurl))
End Sub
