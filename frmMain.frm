VERSION 5.00
Object = "*\AprjFilmStrip.vbp"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   4035
   ClientTop       =   1620
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   10920
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   3000
      ScaleHeight     =   2115
      ScaleWidth      =   4875
      TabIndex        =   4
      Top             =   120
      Width           =   4935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Hidden          =   -1  'True
      Left            =   2880
      Pattern         =   "*.jpg;*.bmp;*.gif"
      System          =   -1  'True
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin prjFilmStrip.UserControl1 UserControl11 
      Height          =   1785
      Left            =   0
      TabIndex        =   0
      Top             =   2640
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   3149
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
File1.Path = Dir1.Path
If File1.ListCount = 0 Then Exit Sub
UserControl11.Clear
For i = 0 To File1.ListCount - 1
    UserControl11.AddItem File1.Path & "\" & File1.List(i)
Next i
DoEvents
UserControl11.Refresh
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then Exit Sub
UserControl11.Width = Me.ScaleWidth
UserControl11.Top = Me.ScaleHeight - UserControl11.Height
End Sub

Private Sub UserControl11_Click(Index As Integer)
Picture1.Picture = LoadPicture(File1.Path & "\" & File1.List(Index))
End Sub

