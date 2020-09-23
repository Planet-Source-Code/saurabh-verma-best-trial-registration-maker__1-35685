VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Main"
   ClientHeight    =   1005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   ScaleHeight     =   1005
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Register"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Try It !"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "You have used 0 days of your 30 day Trial"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim abd As Integer
Dim Registered As Boolean
Dim jj As Integer
Dim st, en As Date

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
Form3.Show
End Sub

Private Sub Command4_Click()
Unload Me
Form2.Show
End Sub

Private Sub Form_Load()
appName = "Microsoft"
secName = "zx12Win"


If GetSetting(appName, secName, "st") <> "×" Then
SaveSetting appName, secName, "st", "×"
SaveSetting appName, secName, "start", Date
SaveSetting appName, secName, "now", Date
SaveSetting appName, secName, "reg", "1"
SaveSetting appName, secName, "alt", "Ö"
End If

If GetSetting(appName, secName, "reg") = "Þ" Then
Unload Me
MsgBox "Software registered"
Form3.Show
Else

st = GetSetting(appName, secName, "start")
en = GetSetting(appName, secName, "now")
abd = DateDiff("d", st, Date)
jj = DateDiff("d", en, Date)

If abd >= 0 And jj >= 0 And GetSetting(appName, secName, "alt") = "Ö" Then
Label1.Caption = "Your " & (30 - abd) & " days left for the try"
Else
SaveSetting appName, secName, "alt", "1"
MsgBox "Sorry !you altered the date "
Unload Me
Form2.Show
End If

End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If GetSetting(appName, secName, "alt") = "Ö" Then
    Dim tt As String
    tt = Date
    SaveSetting appName, secName, "now", tt
End If
End Sub
