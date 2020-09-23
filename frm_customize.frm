VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_customize 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Customize"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_change 
      Caption         =   "Change"
      Height          =   315
      Left            =   5700
      TabIndex        =   2
      Top             =   780
      Width           =   735
   End
   Begin VB.CommandButton cmd_close 
      Caption         =   "Close"
      Height          =   315
      Left            =   6540
      TabIndex        =   1
      Top             =   780
      Width           =   735
   End
   Begin VB.Frame fra_people 
      Caption         =   "People"
      Height          =   735
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7275
      Begin VB.TextBox txt_bitmap 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   6255
      End
      Begin VB.CommandButton cmd_bitmap 
         Caption         =   "Bitmap"
         Height          =   315
         Left            =   6420
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin MSComDlg.CommonDialog com_dialog 
         Left            =   2940
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frm_customize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_bitmap_Click()
    com_dialog.Filter = "Bitmap images (*.bmp)|*.bmp|"
    com_dialog.InitDir = App.Path
    com_dialog.ShowOpen
    If com_dialog.FileName <> "" Then
        txt_bitmap = com_dialog.FileName
    End If
End Sub

Private Sub cmd_change_Click()
    Dim keep As Integer
    Dim previous As String
    previous = GetSetting(App.ProductName, "Customize", "Skin")
    Unload frm_main
    reshaping = True
    reshape_map = txt_bitmap
    Load frm_main
    keep = MsgBox("Do you wish to keep the current skin?", vbYesNo, "Skin change complete")
    If keep = vbYes Then
        SaveSetting App.ProductName, "Customize", "Skin", txt_bitmap
    Else
        Unload frm_main
        reshaping = True
        reshape_map = previous
        Load frm_main
    End If
End Sub

Private Sub cmd_close_Click()
    Unload Me
End Sub

