VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMazeColors 
   Caption         =   "Color Schemes"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   3825
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReset 
      Cancel          =   -1  'True
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "C&ancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   17
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdCreateIt 
      Caption         =   "&Save"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   16
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Color Schemes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3000
         Top             =   2760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdSetSolution 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   15
         Top             =   2880
         Width           =   255
      End
      Begin VB.CommandButton cmdSetCurrent 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton cmdSetVisited 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton cmdSetBackground 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton cmdSetWallColor 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblCLRSolution 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Solution Path:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label lblCLRCurrent 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Current Position:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lblCLRVisited 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Visited Path:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblCLRBackground 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Background:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblCLRWalls 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblWallColor 
         Caption         =   "Wall Color:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmMazeColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCreateIt_Click()
    
    'Update the maze color scheme
    myMaze.clrBackground = lblCLRBackground.BackColor
    myMaze.clrCurrent = lblCLRCurrent.BackColor
    myMaze.clrSolution = lblCLRSolution.BackColor
    myMaze.clrVisited = lblCLRVisited.BackColor
    myMaze.clrWalls = lblCLRWalls.BackColor
    
    'Save the settings to the registry
    SaveSetting "MazeMaker", "Color", "Background", Str(myMaze.clrBackground)
    SaveSetting "MazeMaker", "Color", "Current", Str(myMaze.clrCurrent)
    SaveSetting "MazeMaker", "Color", "Solution", Str(myMaze.clrSolution)
    SaveSetting "MazeMaker", "Color", "Visited", Str(myMaze.clrVisited)
    SaveSetting "MazeMaker", "Color", "Walls", Str(myMaze.clrWalls)
    
    'Close the form
    Unload Me
    
    'Apply the settings
    ApplyColorScheme

End Sub

Private Sub cmdSetBackground_Click()
    'Change the background color
    CommonDialog1.Color = lblCLRBackground.BackColor
    CommonDialog1.ShowColor
    lblCLRBackground.BackColor = CommonDialog1.Color
End Sub

Private Sub cmdSetCurrent_Click()
    'Change the current position color
    CommonDialog1.Color = lblCLRCurrent.BackColor
    CommonDialog1.ShowColor
    lblCLRCurrent.BackColor = CommonDialog1.Color
End Sub

Private Sub cmdSetSolution_Click()
    'Change the solution path color
    CommonDialog1.Color = lblCLRSolution.BackColor
    CommonDialog1.ShowColor
    lblCLRSolution.BackColor = CommonDialog1.Color
End Sub

Private Sub cmdSetVisited_Click()
    'Change the visited path color
    CommonDialog1.Color = lblCLRVisited.BackColor
    CommonDialog1.ShowColor
    lblCLRVisited.BackColor = CommonDialog1.Color
End Sub

Private Sub cmdSetWallColor_Click()
    'Change the wall color
    CommonDialog1.Color = lblCLRWalls.BackColor
    CommonDialog1.ShowColor
    lblCLRWalls.BackColor = CommonDialog1.Color
End Sub

Private Sub Form_Load()
    'Set the initial color values to the color scheme
    lblCLRBackground.BackColor = myMaze.clrBackground
    lblCLRCurrent.BackColor = myMaze.clrCurrent
    lblCLRSolution.BackColor = myMaze.clrSolution
    lblCLRVisited.BackColor = myMaze.clrVisited
    lblCLRWalls.BackColor = myMaze.clrWalls
    
End Sub
