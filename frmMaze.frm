VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMaze 
   Caption         =   "Escape from the Maze!"
   ClientHeight    =   7590
   ClientLeft      =   -180
   ClientTop       =   510
   ClientWidth     =   7260
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   7260
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   7335
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3704
            MinWidth        =   3704
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3704
            MinWidth        =   3704
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4815
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1680
      Top             =   120
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   0
      ScaleHeight     =   465
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   481
      TabIndex        =   0
      Top             =   0
      Width           =   7215
   End
   Begin VB.Menu m1Maze 
      Caption         =   "&Maze"
      Begin VB.Menu m2NewMaze 
         Caption         =   "&New Maze"
      End
      Begin VB.Menu m2ColorScheme 
         Caption         =   "&Change Color Scheme"
      End
      Begin VB.Menu m2ShowSolution 
         Caption         =   "&Show Solution"
      End
      Begin VB.Menu m2Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu m1Help 
      Caption         =   "&Help"
      Begin VB.Menu m2About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMaze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' The display module for the maze
'

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            MoveDirection NORTH
        Case vbKeyDown
            MoveDirection SOUTH
        Case vbKeyLeft
            MoveDirection WEST
        Case vbKeyRight
            MoveDirection EAST
    End Select
    StatusBar1.Panels(1).Text = myMaze.NumberOfMoves & " Moves"
    StatusBar1.Panels(2).Text = myMaze.SolutionMoves & " Moves in Solution"

End Sub

Private Sub Form_Paint()
    DrawMaze
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    'resize the picture window to match the form size, and still remain square
    'wont quite look right if using non-standard maze size (ie 10x60)
    If Me.Width > Me.Height Then
        pic.Height = Me.Height - (2 * StatusBar1.Height) - 750
        pic.Width = pic.Height
        pic.Left = (Me.Width - pic.Width) / 2
    Else
        pic.Width = Me.Width - 400
        pic.Height = pic.Width
        pic.Left = 0
    End If
    ResizeMaze
End Sub

Private Sub m2About_Click()
    MsgBox "MazeMaker 1.0 by Ben Whitney" & vbCrLf & vbCrLf & "This source code is released to the public domain and is available on Planet-Source-Code.com! Please vote for my code and let me know if you like it." & vbCrLf & vbCrLf & "Comments, feedback, questions? Email: benwhitney78@hotmail.com", vbInformation, "About..."
End Sub

Private Sub m2ColorScheme_Click()
    frmMazeColors.Show 1
End Sub

Private Sub m2Exit_Click()
    'End the program
    End
End Sub

Private Sub m2NewMaze_Click()
    myMaze.Initialized = False
    myMaze.OutputPictureBox.Cls
    frmNewMaze.Show 1
End Sub

Private Sub m2ShowSolution_Click()
    ShowSolution
End Sub

Private Sub Timer1_Timer()
    DrawMaze
End Sub
