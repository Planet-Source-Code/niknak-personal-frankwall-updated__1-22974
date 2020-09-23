VERSION 5.00
Begin VB.Form frm_main 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4245
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   510
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tim_check 
      Interval        =   5000
      Left            =   3780
      Top             =   60
   End
   Begin VB.Image img_close 
      Height          =   255
      Left            =   3900
      MouseIcon       =   "frm_main.frx":0CCA
      MousePointer    =   99  'Custom
      Picture         =   "frm_main.frx":19CC
      Stretch         =   -1  'True
      ToolTipText     =   "Double click to quit"
      Top             =   120
      Width           =   255
   End
   Begin VB.Image img_frank 
      Height          =   375
      Left            =   0
      MouseIcon       =   "frm_main.frx":416E
      MousePointer    =   99  'Custom
      Top             =   60
      Width           =   315
   End
   Begin VB.Label lbl_alert 
      Alignment       =   2  'Center
      Caption         =   "Cant!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Image img_shapemap 
      Height          =   510
      Left            =   0
      Picture         =   "frm_main.frx":4E70
      Top             =   0
      Width           =   4245
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Type POINTAPI
        x As Long
        y As Long
End Type

Dim hidden As Boolean
Dim firstload As Boolean

'LOAD AND SHAPE THE FORM
Private Sub Form_Load()
    If Not firstload Then
        check_ipadd
        If GetSetting(App.ProductName, "Customize", "Skin") <> "" Then
            reshaping = True
            reshape_map = GetSetting(App.ProductName, "Customize", "Skin")
        End If
        firstload = True
    End If
    If reshaping Then
        load_window App.ProductName
        img_shapemap = LoadPicture(reshape_map)
        SavePicture img_shapemap.Picture, App.Path & "\shapemap.tmp"
        Face = CreateRegionFromFile(Me, img_shapemap, App.Path & "\shapemap.tmp", RGB(0, 255, 0))
        SetWindowRgn Me.hwnd, Face, True
        hideme
        Me.Visible = True
        reshaping = False
    Else
        load_window App.ProductName
        SavePicture img_shapemap.Picture, App.Path & "\shapemap.tmp"
        Face = CreateRegionFromFile(Me, img_shapemap, App.Path & "\shapemap.tmp", RGB(0, 255, 0))
        SetWindowRgn Me.hwnd, Face, True
        hideme
        Me.Visible = True
    End If
    Load frm_winsock
    showme "IP Changed to " & frm_winsock.win_frank.LocalIP
    Unload frm_winsock
End Sub

'HIDE FRANKBAR
Private Sub hideme()
    img_frank.ToolTipText = "Double click to open"
    Dim next_pos As Long
    DoEvents
    For next_pos = Me.Left To (Screen.Width - 187) Step 1
        Me.Left = next_pos
        Sleep 0.1
    Next next_pos
    Me.Left = Screen.Width - 187
    hidden = True
End Sub

'SHOW FRANKBAR
Private Sub showme(Optional message As String)
    If message <> "" Then
        hideme
        lbl_alert.Caption = message
    End If
    img_frank.ToolTipText = "Double click to close"
    Dim next_pos As Long
    DoEvents
    For next_pos = Me.Left To (Screen.Width - Me.Width) Step -4
        Me.Left = next_pos
        Me.Refresh
    Next next_pos
    Me.Left = Screen.Width - Me.Width
    hidden = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    save_window App.ProductName, Me.Top
End Sub

Private Sub img_close_DblClick()
    Unload Me
End Sub

'SHOW AND HIDE
Private Sub img_frank_DblClick()
    If hidden Then
        showme
    Else
        hideme
    End If
End Sub

'MOVE FRANKBAR
Private Sub img_frank_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
    Else
        frm_customize.Show
    End If
    If hidden Then
        hideme
    Else
        showme
    End If
End Sub

'SAVE AND LOAD FRANKBAR POSITION
Public Sub save_window(window As String, save_top As Long)
    SaveSetting App.ProductName, "windows", window, "SAVED"
    SaveSetting App.ProductName, "windows", window & " top", save_top
End Sub

Public Sub load_window(window As String)
    If GetSetting(App.ProductName, "windows", window) = "SAVED" Then
        win_top = Val(GetSetting(App.ProductName, "windows", window & " top"))
    Else
        win_top = 0
    End If
    Me.Top = win_top
End Sub

Private Sub tim_check_Timer()
    check_ipadd
End Sub

Private Sub check_ipadd()
    Static current_ipadd As String
    Load frm_winsock
    If frm_winsock.win_frank.LocalIP <> current_ipadd Then
        current_ipadd = frm_winsock.win_frank.LocalIP
        showme "IP Changed to " & current_ipadd
    End If
    Unload frm_winsock
End Sub
