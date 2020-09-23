VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search Options"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   300
      Left            =   2055
      TabIndex        =   3
      Top             =   540
      Width           =   810
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   300
      Left            =   1215
      TabIndex        =   2
      Top             =   540
      Width           =   810
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   30
      TabIndex        =   1
      Top             =   210
      Width           =   4005
   End
   Begin VB.Label Label1 
      Caption         =   "URL to a search page's search engine:"
      Height          =   285
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   2865
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    SaveSetting "Search", "Options", "URL", Text1.Text
    Beep
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Text1.Text = GetSetting("Search", "Options", "URL", "http://www.altavista.com/cgi-bin/query?pg=q&q=")
End Sub
