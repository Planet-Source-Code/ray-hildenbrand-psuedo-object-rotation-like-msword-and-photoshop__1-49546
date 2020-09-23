VERSION 5.00
Begin VB.Form frmTestRotation 
   Caption         =   "Form1"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Set RenderedPicture"
      Enabled         =   0   'False
      Height          =   240
      Left            =   5235
      TabIndex        =   7
      Top             =   210
      Width           =   1785
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Rotate 270"
      Enabled         =   0   'False
      Height          =   510
      Left            =   4755
      TabIndex        =   6
      Top             =   3585
      Width           =   1410
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Rotate 180"
      Enabled         =   0   'False
      Height          =   510
      Left            =   2940
      TabIndex        =   5
      Top             =   3570
      Width           =   1410
   End
   Begin VB.PictureBox picRendered 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3105
      Left            =   3660
      ScaleHeight     =   3075
      ScaleWidth      =   3495
      TabIndex        =   2
      Top             =   300
      Width           =   3525
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Rotate 90"
      Enabled         =   0   'False
      Height          =   510
      Left            =   1140
      TabIndex        =   1
      Top             =   3540
      Width           =   1410
   End
   Begin VI_Rotator.ImageRotation orTest 
      Height          =   2520
      Left            =   210
      TabIndex        =   8
      Top             =   450
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   4445
   End
   Begin VB.PictureBox picBG 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   75
      ScaleHeight     =   3105
      ScaleWidth      =   3285
      TabIndex        =   0
      Top             =   285
      Width           =   3315
   End
   Begin VB.Label Label2 
      Caption         =   "Rendered Picture"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3660
      TabIndex        =   4
      Top             =   75
      Width           =   2100
   End
   Begin VB.Label Label1 
      Caption         =   "Rotation Workspace"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   75
      TabIndex        =   3
      Top             =   75
      Width           =   2100
   End
   Begin VB.Image imgImage 
      Height          =   1920
      Left            =   795
      Picture         =   "frmTestRotation.frx":0000
      Top             =   300
      Width           =   1920
   End
End
Attribute VB_Name = "frmTestRotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim beeninit  As Boolean
Private Sub Command2_Click()
    orTest.DrawImageAtAngle 90, True
End Sub

Private Sub Command3_Click()
orTest.DrawImageAtAngle 180, True
End Sub

Private Sub Command4_Click()
orTest.DrawImageAtAngle 270, True
End Sub

Private Sub Command5_Click()
    Set picRendered.Picture = orTest.RenderedPicture
End Sub

Private Sub Form_Load()
    If Not beeninit Then
        Set orTest.OriginalPicture = imgImage.Picture
        Command2.Enabled = True
        Command3.Enabled = True
        Command4.Enabled = True
        Command5.Enabled = True
    End If
End Sub
