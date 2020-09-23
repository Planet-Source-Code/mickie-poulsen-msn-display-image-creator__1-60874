VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form frmPartialScreenShot 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MSN SmileyCreator 1.2"
   ClientHeight    =   8415
   ClientLeft      =   5250
   ClientTop       =   795
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   561
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   478
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5760
      TabIndex        =   11
      Text            =   "FFFFFF"
      Top             =   6600
      Width           =   1335
   End
   Begin VB.PictureBox picSave 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   225
      ScaleHeight     =   1440
      ScaleWidth      =   1440
      TabIndex        =   13
      Top             =   6705
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Full Preview"
      Height          =   360
      Left            =   2640
      TabIndex        =   15
      Top             =   9360
      Width           =   1500
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Image"
      Height          =   360
      Left            =   4200
      TabIndex        =   14
      Top             =   9360
      Width           =   1500
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6720
      Top             =   7440
   End
   Begin VB.PictureBox picScreenShot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   225
      ScaleHeight     =   1440
      ScaleWidth      =   1440
      TabIndex        =   9
      Top             =   6705
      Width           =   1440
   End
   Begin VB.TextBox txtHeight 
      Height          =   288
      Left            =   11400
      TabIndex        =   7
      Text            =   "96"
      Top             =   10200
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtWidth 
      Height          =   288
      Left            =   11400
      TabIndex        =   5
      Text            =   "96"
      Top             =   9840
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtLeft 
      Height          =   288
      Left            =   9600
      TabIndex        =   3
      Top             =   10200
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtTop 
      Height          =   288
      Left            =   9600
      TabIndex        =   1
      Top             =   9840
      Visible         =   0   'False
      Width           =   732
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash SWF1 
      Height          =   5295
      Left            =   -120
      TabIndex        =   10
      Top             =   1200
      Width           =   7290
      _cx             =   12859
      _cy             =   9340
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   "FFFFFF"
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "By Mickie Poulsen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   180
      TabIndex        =   20
      Top             =   900
      Width           =   6795
   End
   Begin VB.Line Line2 
      X1              =   -16
      X2              =   480
      Y1              =   79
      Y2              =   79
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "This is your final image, if you are satisfied with the result, please hit the ""Save to .bmp"" button..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2280
      TabIndex        =   19
      Top             =   6720
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Save to .bmp"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   7995
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Full Preview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   17
      Top             =   7995
      Width           =   1335
   End
   Begin VB.Image Image5 
      Height          =   375
      Left            =   4080
      Picture         =   "frmPartialScreenShot.frx":0000
      Top             =   7920
      Width           =   1665
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   2280
      Picture         =   "frmPartialScreenShot.frx":05FA
      Top             =   7920
      Width           =   1665
   End
   Begin VB.Shape Shape1 
      Height          =   1470
      Left            =   210
      Top             =   6690
      Width           =   1470
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   480
      Y1              =   433
      Y2              =   433
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MY CAPTION"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2955
      TabIndex        =   16
      Top             =   3450
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Set BG Color (Web)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00303030&
      Height          =   255
      Left            =   5745
      TabIndex        =   12
      Top             =   6930
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   1215
      Left            =   -120
      Picture         =   "frmPartialScreenShot.frx":0BF4
      Top             =   0
      Width           =   7290
   End
   Begin VB.Image frame1 
      Appearance      =   0  'Flat
      Height          =   1440
      Left            =   2850
      Top             =   2280
      Width           =   1440
   End
   Begin VB.Image Image1 
      Height          =   5295
      Left            =   -120
      Top             =   1200
      Width           =   7290
   End
   Begin VB.Label lblNote 
      Caption         =   "N.B. These values are in pixels."
      Height          =   495
      Left            =   12360
      TabIndex        =   8
      Top             =   9960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblTop 
      Caption         =   "&Top"
      Height          =   255
      Left            =   8880
      TabIndex        =   0
      Top             =   9840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblWidth 
      Caption         =   "&Width"
      Height          =   255
      Left            =   10680
      TabIndex        =   4
      Top             =   9840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblHeight 
      Caption         =   "&Height"
      Height          =   255
      Left            =   10680
      TabIndex        =   6
      Top             =   10200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblLeft 
      Caption         =   "&Left"
      Height          =   255
      Left            =   8880
      TabIndex        =   2
      Top             =   10200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   3240
      Left            =   -1380
      Picture         =   "frmPartialScreenShot.frx":5AFA
      Top             =   6225
      Width           =   3705
   End
End
Attribute VB_Name = "frmPartialScreenShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdSave_Click()
    SavePicture picSave.Image, App.Path + "\Smilie.bmp"
End Sub

Private Sub Command1_Click()
Form1.Show
Form1.Image1.Picture = frmPartialScreenShot.picSave.Image
End Sub

Private Sub Form_Load()
    txtTop.Text = frmPartialScreenShot.Top
    txtLeft.Text = frmPartialScreenShot.Left
    SWF1.Movie = App.Path & "\creator.swf"
End Sub




Private Sub Image4_Click()
Form1.Show
Form1.Image1.Picture = frmPartialScreenShot.picSave.Image
End Sub

Private Sub Image5_Click()
' stop capturing.
Timer1.Enabled = False
On Error GoTo ErroR

  Dim cDlg As New cDialog
  Dim FileName As String

FileName = cDlg.GetSaveName("GetSaveName API Test")
SavePicture picSave.Image, FileName & ".bmp"


Timer1.Enabled = True
ErroR:
Timer1.Enabled = True
End Sub

Private Sub Label1_Click()
Form1.Show
Form1.Image1.Picture = frmPartialScreenShot.picSave.Image
End Sub

Private Sub Label4_Click()
'' stop capturing.
Timer1.Enabled = False
On Error GoTo ErroR



  Dim cDlg As New cDialog
  Dim FileName As String
  


FileName = cDlg.GetSaveName("GetSaveName API Test")

If FileName = "" Then GoTo ErroR
SavePicture picSave.Image, FileName & ".bmp"


Timer1.Enabled = True
ErroR:
Timer1.Enabled = True
End Sub

Private Sub Text1_Change()
On Error Resume Next
    SWF1.BGColor = Text1.Text
End Sub

Private Sub Timer1_Timer()
Dim Top As Long, Left As Long, Width As Long, Height As Long
On Error GoTo Err 'in case the .Text are not numeric

    txtTop.Text = (Screen.Width / Screen.TwipsPerPixelX / 2) - (frmPartialScreenShot.ScaleWidth / 2) + frame1.Left - 1
    txtLeft.Text = (Screen.Height / Screen.TwipsPerPixelY / 2) - (frmPartialScreenShot.ScaleHeight / 2) + frame1.Top + 12
    Top = CLng(txtTop.Text)
    Left = CLng(txtLeft.Text)
    Width = CLng(txtWidth.Text)
    Height = CLng(txtHeight.Text)
    Call PartialScreenShot(picScreenShot.hdc, picSave, Top, Left, Width, Height)
    
Err:
End Sub



'' A HUGE THANK YOU TO WILKSEY FOR FIXING THE SAVE PROBLEM!!
'' PLEASE CHECK SOME OF HIS WORK AT http://www.pscode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&txtCriteria=Wilksey&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&lngWId=1&B1=Quick+Search
'' ;)
