VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_browse 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "i2"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12180
   Icon            =   "frm_browse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   12180
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar status 
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   4650
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   30
      Pattern         =   "*.jpg"
      TabIndex        =   2
      Top             =   2310
      Width           =   2475
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   30
      TabIndex        =   1
      Top             =   390
      Width           =   2475
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   2475
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   2400
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label lbl_about 
      Caption         =   "ABOUT"
      Height          =   1275
      Left            =   60
      TabIndex        =   5
      Top             =   5970
      Width           =   2385
   End
   Begin VB.Image mydocuments 
      Height          =   795
      Left            =   60
      Picture         =   "frm_browse.frx":0442
      Top             =   5010
      Width           =   1125
   End
   Begin VB.Image mypictures 
      Height          =   795
      Left            =   1290
      Picture         =   "frm_browse.frx":33B8
      Top             =   5010
      Width           =   1125
   End
   Begin VB.Label lbl_filename 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Image Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   690
      TabIndex        =   3
      Top             =   4380
      Width           =   1095
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   7200
      Left            =   2550
      Picture         =   "frm_browse.frx":632E
      Stretch         =   -1  'True
      Top             =   30
      Width           =   9600
   End
End
Attribute VB_Name = "frm_browse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
   File1.Path = Dir1.Path   ' Set file path.
End Sub

Private Sub Drive1_Change()
On Error GoTo err
Dir1.Path = Drive1.Drive
lbl_filename.Caption = ""
Exit Sub
err:
    MsgBox "The device could not be opened.", vbCritical, "Error"
    Exit Sub
End Sub

Private Sub File1_Click()
If File1.ListIndex = -1 Then Exit Sub Else _
status.Value = 0
lbl_filename.Caption = File1.FileName
'**************************************************
picpath = File1.Path & "\" & lbl_filename.Caption
Set Image1.Picture = LoadPicture(picpath)
Me.Caption = "i2 - " & lbl_filename.Caption
status.Value = 100
End Sub

Private Sub Form_Load()
lbl_about.Caption = "i2 - Copyright iawix.com 2000" & vbCrLf & _
"Version " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
"Please visit iawix.com for my other software and updates to i2."
GoTo 1
1:
On Error GoTo err1
File1.Path = "C:\My Documents"
GoTo 2
err1:
'Must be win95, no my documents
mydocuments.Enabled = False
mypictures.Enabled = False
Exit Sub
2:
On Error GoTo err2
File1.Path = "C:\My Documents\My Pictures"
GoTo 3
err2:
'Must be win 98, no my pictures (WIN ME)
mypictures.Enabled = False
Exit Sub
3:
'WIN ME
mydocuments.Enabled = True
mypictures.Enabled = True
End Sub

Private Sub mypictures_Click()
Dir1.Path = "C:\My Documents\My Pictures"
File1.Path = "C:\My Documents\My Pictures"
End Sub

Private Sub mydocuments_Click()
Dir1.Path = "C:\My Documents"
File1.Path = "C:\My Documents"
End Sub
