VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNewMaze 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MazeMaker 1.0 by Ben Whitney"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5190
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4800
      Top             =   2040
   End
   Begin VB.CommandButton cmdCreateIt 
      Caption         =   "&Create It!"
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
      Left            =   2880
      TabIndex        =   10
      Top             =   4620
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "C&lose"
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
      Left            =   4560
      TabIndex        =   9
      Top             =   4620
      Width           =   1455
   End
   Begin VB.Frame frameNewMaze 
      Caption         =   "New Maze"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   6015
      Begin MSComctlLib.Slider sldrComplexity 
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2640
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   4
         SelStart        =   2
         Value           =   2
      End
      Begin VB.ComboBox cboCols 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmNewMaze.frx":0000
         Left            =   2760
         List            =   "frmNewMaze.frx":0016
         TabIndex        =   8
         Text            =   "50"
         Top             =   1680
         Width           =   735
      End
      Begin VB.ComboBox cboRows 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmNewMaze.frx":0032
         Left            =   1680
         List            =   "frmNewMaze.frx":0048
         TabIndex        =   6
         Text            =   "50"
         Top             =   1680
         Width           =   735
      End
      Begin VB.OptionButton opDifficulty 
         Caption         =   "Custom Size:"
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
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   1740
         Width           =   1335
      End
      Begin VB.OptionButton opDifficulty 
         Caption         =   "Hard (40 x 40)"
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
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   3255
      End
      Begin VB.OptionButton opDifficulty 
         Caption         =   "Medium (20 x 20)"
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
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   1155
         Width           =   3255
      End
      Begin VB.OptionButton opDifficulty 
         Caption         =   "Easy (10 x 10)"
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
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label lblComplexity 
         Caption         =   "Normal"
         Height          =   255
         Left            =   2520
         TabIndex        =   15
         Top             =   2640
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "Maze Complexity"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblBy 
         Caption         =   "x"
         Enabled         =   0   'False
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
         Left            =   2520
         TabIndex        =   7
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label lblMazeDifficulty 
         Caption         =   "Maze Difficulty"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1800
      TabIndex        =   12
      Top             =   840
      Width           =   1245
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   0
      Picture         =   "frmNewMaze.frx":0064
      Top             =   0
      Width           =   6315
   End
   Begin VB.Label lblBlack 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   840
      Width           =   6375
   End
End
Attribute VB_Name = "frmNewMaze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'' ------------------------------------------------------------
'' FILE:            frmNewMaze.frm
'' AUTHOR:          Ben Whitney (benwhitney78@hotmail.com)
'' DESCRIPTION:     This new form will initialize and create
''                  a new maze.  Any parameters changed will
''                  saved to the registry, so that this form,
''                  when re-displayed, will be able to maintain
''                  its previous state.
'' ------------------------------------------------------------
Private mnDifficultyLevel As Integer
Private Sub cboRows_Validate(Cancel As Boolean)
    If Not IsNumeric(cboRows.Text) Then
        MsgBox "Please enter a valid size (" & MIN_SIZE & " to " & MAX_SIZE & ") "
        Cancel = True
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    End
End Sub

Private Sub cmdCreateIt_Click()
    Dim nRows As Integer
    Dim nCols As Integer
    Dim nComplex As Double
    
    'Save all the settings, then create the maze
    SaveSetting "MazeMaker", "Init", "LastChoice", Str(mnDifficultyLevel)
    
    'set the number of rows and colums based on the option chosen
    Select Case mnDifficultyLevel
        Case 0
            nRows = 10
            nCols = 10
        Case 1
            nRows = 20
            nCols = 20
        Case 2
            nRows = 40
            nCols = 40
        Case 3
            'Only save the custom maze settings if that option was chosen
            nRows = Val(cboRows.Text)
            nCols = Val(cboCols.Text)
            SaveSetting "MazeMaker", "Init", "LastRows", cboRows.Text
            SaveSetting "MazeMaker", "Init", "LastCols", cboCols.Text
    End Select
    
    nComplex = 1 + (0.5 * sldrComplexity.Value)
    
    InitializeMaze nRows, nCols, frmMaze.pic, nComplex
    
    Unload Me
    frmMaze.Show
    StartMaze
End Sub

Private Sub Form_Load()
    Initialize
End Sub

Private Sub Initialize()
    'All initialization routines occur here.
    
    On Local Error GoTo Initialize_Validation_Error
    
    'Reminder:
    'Const MIN_SIZE = 5               'Define minimum and maximum size
    'Const MAX_SIZE = 100             'for validation
    
    Dim nLastNumberOfRows As Integer 'Last number of rows that were selected in custom mode
    Dim nLastNumberOfCols As Integer 'Last number of columns that were selected in custom mode
    Dim nLastChoice As Integer       'Previous choice for type of maze, corresponds to index of opDifficulty control
    Dim sTemp As String              'Temporary string variable for reading registry settings
        
    'Read in previous variables from the registry
    sTemp = GetSetting("MazeMaker", "Init", "LastChoice", "1")
    
    'Validate the data for the last choice selected
    If IsNumeric(sTemp) Then
        nLastChoice = Int(Val(sTemp))
        If nLastChoice < 0 Or nLastChoice > 3 Then
            'invalid choice range, set to zero
            nLastChoice = 0
        End If
    Else
        nLastChoice = 0
    End If
    
    'Validate the data for the last number of rows
    sTemp = GetSetting("MazeMaker", "Init", "LastRows", "50")
    
    If IsNumeric(sTemp) Then
        nLastNumberOfRows = Int(Val(sTemp))
        If nLastNumberOfRows < MIN_SIZE Then
            nLastNumberOfRows = MIN_SIZE
        ElseIf nLastNumberOfRows > MAX_SIZE Then
            nLastNumberOfRows = MAX_SIZE
        End If
    Else
        nLastNumberOfRows = 50
    End If
    
    'Validate the data for the last number of columns
    sTemp = GetSetting("MazeMaker", "Init", "LastCols", "50")
    
    If IsNumeric(sTemp) Then
        nLastNumberOfCols = Int(Val(sTemp))
        If nLastNumberOfCols < MIN_SIZE Then
            nLastNumberOfCols = MIN_SIZE
        ElseIf nLastNumberOfCols > MAX_SIZE Then
            nLastNumberOfCols = MAX_SIZE
        End If
    Else
        nLastNumberOfCols = 50
    End If
    
    'Set the appropriate choices
    opDifficulty(nLastChoice).Value = True
    SetComboValue cboRows, Str(nLastNumberOfRows)
    SetComboValue cboCols, Str(nLastNumberOfCols)
    
    Exit Sub

'Error handling code

Initialize_Validation_Error:
    'handle for the registry reading section
    Debug.Print ">> Initialize_Validation_Error (" & Str(Err.Number) & ") " & Err.Description
    Resume Next
End Sub

Public Sub SetComboValue(combo As ComboBox, sValue As String)
    'This function will set the value of a combo box to the specified
    'value, and select the appropriate index if the value is contained
    'in its drop-down list

    Dim x As Integer                'loop variable
    Dim bFound As Boolean           'found value in list, or not
    
    On Local Error GoTo SetComboValue_Error
    
    If combo.ListCount > 0 Then     'ensure there is at least 1 item to check agains
        
        For x = 0 To combo.ListCount - 1
            If Trim(combo.List(x)) = Trim(sValue) Then
                combo.ListIndex = x
                bFound = True
                Exit For
            End If
        Next x
                
        If Not bFound Then          'value was not found in the drop down list
            combo.Text = sValue     'set text property instead
        End If
    Else
        combo.Text = sValue         'no drop-down list, set text property instead
    End If
    
    Exit Sub
    
    'error handling code

SetComboValue_Error:
    Debug.Print ">> SetComboValue_Error (" & Str(Err.Number) & ") " & Err.Description
    Resume Next
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        'the x close button was pressed
        End
    End If
End Sub

Private Sub opDifficulty_Click(Index As Integer)
    If Index = 3 Then
        'Custom mode has been selected, enable the custom controls
        cboRows.Enabled = True
        cboCols.Enabled = True
        lblBy.Enabled = True
        If frmNewMaze.Visible = True Then
            cboRows.SetFocus
            SendKeys "{F4}"     'trigger a dropdown
        End If
    Else
        cboRows.Enabled = False
        cboCols.Enabled = False
        lblBy.Enabled = False
    End If
    mnDifficultyLevel = Index
End Sub

Private Sub sldrComplexity_Change()
    Select Case sldrComplexity
        Case 1
            lblComplexity = "Low"
        Case 2
            lblComplexity = "Normal"
        Case 3
            lblComplexity = "High"
        Case 4
            lblComplexity = "Extreme (May take a long time)"
    End Select
End Sub


Private Sub Timer1_Timer()
    'Color cycle timer
    
    Static nClr As Integer          'the current color value, 0-255
    Static nDirection As Integer    '+1 or -1, depending on which way
                                    'the color is heading from white
    
    Dim msgs(8) As String
    msgs(1) = "Hello?"
    msgs(2) = "Can anybody hear me?"
    msgs(3) = "I'm lost."
    msgs(4) = "It's dark in here."
    msgs(5) = "It's dark and I'm frightened..."
    msgs(6) = "Who turned out the lights?"
    msgs(7) = "Can somebody please help me?"
    msgs(8) = "I am so lost, its not even funny."
    
    If nDirection = 0 Then          'first time initializtion
        nClr = 255
        nDirection = -1
        lblFind = msgs(6)
        Randomize Timer
    End If
    
    If nClr > 255 Then
        nClr = 255
        nDirection = -nDirection    'switch directions
    ElseIf nClr < 0 Then
        nClr = 0
        nDirection = -nDirection
        'move the question to a new position
        lblFind.Caption = msgs(Int(Rnd * 8))
        lblFind.Top = lblBlack.Top + Int(Rnd * (lblBlack.Height - lblFind.Height))
        lblFind.Left = Int(Rnd * (lblBlack.Width - lblFind.Width))
        
    End If
    
    lblFind.ForeColor = RGB(nClr, nClr, nClr)
    
    nClr = nClr + nDirection
    
    
End Sub
