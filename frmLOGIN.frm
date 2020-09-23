VERSION 5.00
Begin VB.Form frmLOGIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGIN"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   3525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame FrNewInstall 
      Caption         =   "Is this a new installation?"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.OptionButton OptionNeither 
         Caption         =   "Neither"
         Height          =   495
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton OptionNo 
         Caption         =   "No"
         Height          =   495
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton optionYes 
         Caption         =   "Yes"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmLOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdContinue_Click()
If frmLOGIN.optionYes.Value = True Then
SaveSetting "EULA", "SETTINGS", "NewYes", frmLOGIN.optionYes.Value
frmEULA.Show
frmLOGIN.Hide
Else
End If
End Sub
Private Sub OptionNo_Click()
'If EULA has been previously agreed upon, skips over EULA form
'to Splash page.
'User won't go any further until agreeing to EULA terms
'OptionNeither button is a dummy set to TRUE to catch
'the initial focus.
'Continue Text button is hidden off of form. BorderStyle property
'is Fixed Single
Dim RegistryKey As String
On Error GoTo errhandler
frmEULA.OBEULA1.Value = GetSetting("EULA", "SETTINGS", "NewYes")

If frmEULA.OBEULA1.Value = False Then
frmEULA.Hide
frmLOGIN.Hide
frmSplash.Show
Else
End
End If
errhandler: 'dummy block
End Sub

Private Sub optionYes_Click()
If frmLOGIN.optionYes.Value = True Then
SaveSetting "EULA", "SETTINGS", "NewYes", frmLOGIN.optionYes.Value
frmEULA.Show
frmLOGIN.Hide
Else
End If
End Sub
