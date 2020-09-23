VERSION 5.00
Begin VB.Form search 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   345
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   6870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   345
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton search 
      Caption         =   "Search"
      Height          =   315
      Left            =   5565
      TabIndex        =   2
      Top             =   15
      Width           =   1290
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   30
      TabIndex        =   0
      Text            =   "Enter search string here"
      Top             =   30
      Width           =   5490
   End
   Begin VB.Image searchicon 
      Height          =   240
      Left            =   30
      Picture         =   "search.frx":0000
      Top             =   465
      Width           =   240
   End
   Begin VB.Label move 
      Height          =   420
      Left            =   -225
      TabIndex        =   1
      Top             =   -30
      Width           =   7455
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp_Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu mnuline1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowSearchWindow 
         Caption         =   "&Show search window"
      End
      Begin VB.Menu mnuline2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MouseStartX As Long
Public MouseStartY As Long

Private Sub Form_Load()
    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.hWnd = Me.hWnd
    TrayIcon.uId = vbNull
    TrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayIcon.ucallbackMessage = WM_MOUSEMOVE
    TrayIcon.hIcon = searchicon.Picture
    TrayIcon.szTip = "Mind's Tray Icon Example" & Chr$(0)
    Call Shell_NotifyIcon(NIM_ADD, TrayIcon)
End Sub

Private Sub mnuExit_Click()
    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.hWnd = Me.hWnd
    TrayIcon.uId = vbNull
    Call Shell_NotifyIcon(NIM_DELETE, TrayIcon)
    End
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show
End Sub

Private Sub move_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseStartX = X
    MouseStartY = Y
End Sub

Private Sub move_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbLeftButton Then
        Me.Left = Me.Left - (MouseStartX - X)
        Me.Top = Me.Top - (MouseStartY - Y)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Static Message As Long
Static RR As Boolean
    Message = X / Screen.TwipsPerPixelX
    
    If RR = False Then
        RR = True
        Select Case Message
            Case WM_LBUTTONDBLCLK
                Me.Show
            Case WM_RBUTTONUP
                Me.PopupMenu mnuPopUp
        End Select
        RR = False
    End If
    
End Sub

Private Sub search_Click()
    If Text1.Text = "" Then
        MsgBox "You must enter a search string!", vbOKOnly + vbWarning, "Search"
    Else
        searchurl = GetSetting("Search", "Options", "URL", "http://www.altavista.com/cgi-bin/query?pg=q&q=")
        Shell ("start " + searchurl + Text1.Text), vbHide
    End If
End Sub
