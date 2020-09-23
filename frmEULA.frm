VERSION 5.00
Begin VB.Form frmEULA 
   Caption         =   "End User License Agreement"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OBEULA2 
      Caption         =   "Option2"
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   2880
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton OBEULA1 
      Caption         =   "Option1"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdDISAGREE 
      Caption         =   "I Disagree"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdAGREE 
      Caption         =   "I Agree"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtEULA 
      Height          =   2655
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmEULA.frx":0000
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmEULA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''I needed a way to have the User agree to some legal terms,
'''so I developed this. I'm sure there are other ways but
'''this is what I wrote. Writes User decision to the Registry.
'''This code is set up so that once the User agrees to the EULA,
'''the form will skip over the EULA from then on so that it doesn't
'''annoy the User. Replace the items within quotation marks ""
'''on frmEULA txtEULA textbox (in the Text property)
'''to match your needs.
'''This is my first post, so please be kind in your remarks.
'''Vote if you wish, offer suggestions etc...
'''This is just me trying to help people for those who helped me.
'''In the txtEULA text box, fill in all the appropriate data
'''within the quote (" ") marks.

Private Sub cmdAGREE_Click()
OBEULA1.Value = True
SaveSetting "EULA", "SETTINGS", "NewYes", OBEULA2.Value
If OBEULA1.Value = True Then
frmSplash.Show
frmEULA.Hide
Else
frmEULA.Show
End If
End Sub
Private Sub cmdDISAGREE_Click()
OBEULA1.Value = False
SaveSetting "EULA", "SETTINGS", "NewYes", frmLOGIN.optionYes.Value
Unload Me
End
End Sub

Private Sub Form_Activate()
If OBEULA1.Value = False Then
frmEULA.Hide
frmSplash.Show
Else
End If
End Sub

Private Sub Form_Load()
OBEULA1.Value = GetSetting("EULA", "SETTINGS", "NewYes")
End Sub
