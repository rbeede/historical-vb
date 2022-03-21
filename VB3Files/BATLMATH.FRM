VERSION 2.00
Begin Form frmMain 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BattleMath"
   ClientHeight    =   4020
   ClientLeft      =   1230
   ClientTop       =   1635
   ClientWidth     =   7365
   ForeColor       =   &H000000FF&
   Height          =   4425
   Icon            =   BATLMATH.FRX:0000
   Left            =   1170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7365
   Top             =   1290
   Width           =   7485
   Begin PictureBox pctShips 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3525
      Left            =   7440
      Picture         =   BATLMATH.FRX:0302
      ScaleHeight     =   3525
      ScaleWidth      =   2070
      TabIndex        =   8
      Top             =   960
      Width           =   2070
   End
   Begin Timer tmrFlashGrid 
      Interval        =   250
      Left            =   1560
      Top             =   2280
   End
   Begin CommandButton cmdHelp 
      Caption         =   "&Help"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   18
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
   End
   Begin CommandButton cmdResign 
      Caption         =   "&Resign"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   18
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin CommandButton cmdStart 
      Caption         =   "&Start"
      Enabled         =   0   'False
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   18
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin ListBox lstSkills 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin Label lblGrid 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Your Grid"
      Height          =   195
      Left            =   4260
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   825
   End
   Begin Image imgHitBox 
      Height          =   465
      Left            =   4080
      Picture         =   BATLMATH.FRX:87C0
      Top             =   3480
      Visible         =   0   'False
      Width           =   465
   End
   Begin Image imgGridBox 
      Height          =   465
      Left            =   4080
      Picture         =   BATLMATH.FRX:8A2A
      Top             =   3000
      Visible         =   0   'False
      Width           =   465
   End
   Begin Image imgMissBox 
      Height          =   465
      Left            =   3840
      Picture         =   BATLMATH.FRX:8C94
      Top             =   3960
      Visible         =   0   'False
      Width           =   465
   End
   Begin Image imgShiphBox 
      Height          =   465
      Left            =   3600
      Picture         =   BATLMATH.FRX:8EFE
      Top             =   3480
      Visible         =   0   'False
      Width           =   465
   End
   Begin Image imgShipBox 
      Height          =   465
      Left            =   3600
      Picture         =   BATLMATH.FRX:9168
      Top             =   3000
      Visible         =   0   'False
      Width           =   465
   End
   Begin Label lblInstructions 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblInstructions"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   585
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   3210
   End
   Begin Label lblSkill 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Skill:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1845
   End
   Begin Label lblPCShots 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Shots I get: ##"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   18
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   465
      Left            =   6000
      TabIndex        =   1
      Top             =   5880
      Visible         =   0   'False
      Width           =   2535
   End
   Begin Label lblMyShots 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Shots you get: ##"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   18
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   465
      Left            =   120
      TabIndex        =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   3045
   End
   Begin Image imgOceanGrid 
      Height          =   465
      Index           =   0
      Left            =   2280
      Picture         =   BATLMATH.FRX:93D2
      Top             =   960
      Width           =   465
   End
End
Const ptrDefault = 0 'Variable for mouse pointer default look
Const ptrNS = 7 'Variable for mouse pointer north south look
Const ptrWE = 9 'Variable for mouse pointer west east look

'Declare variables
Dim Updated(5) As Integer 'For cells to update
Dim UpdatedAmt As Integer 'For how many cells to update
Dim MyGrid(99) As String 'Array for storing where ships are
Dim PCGrid(99) As String 'Array for storing computer ships
Dim MyShots(99) As String 'Array for storing where shot at
Dim PCShots(99) As String 'Array for storing where PC shot at
Dim Started As Integer 'Flag to check if started game
Dim Ship As String 'Tells which ship is being placed
Dim Passed As Integer 'Tell if ship placement is valid
Dim DirectionChange As Integer 'Tell if direction has changed
Dim MyHits As Integer 'Keep count of how many ships hit
Dim PCHits As Integer 'Keep count of how many ships hit

Function CheckPlacement (Index As Integer) As Integer

   'Declare variables
   Dim X As Integer 'For counter
   Dim Border As Integer 'For calculation of border
   
   On Error Resume Next 'Skip errors
   
   Passed = True 'Assume passed
   
   'Check for where ship is being placed
   'If invalid placement change flag

   If Screen.MousePointer = ptrNS Then 'Cursor is N S
      Select Case Ship 'Determine which ship to check
         Case "Carrier" 'Carrier ship
             If MyGrid(Index - 20) = "Ship" Then Passed = False
             If MyGrid(Index - 10) = "Ship" Then Passed = False
             If MyGrid(Index) = "Ship" Then Passed = False
             If MyGrid(Index + 10) = "Ship" Then Passed = False
             If MyGrid(Index + 20) = "Ship" Then Passed = False

          Case "Battleship" 'Battleship
             If MyGrid(Index - 10) = "Ship" Then Passed = False
             If MyGrid(Index) = "Ship" Then Passed = False
             If MyGrid(Index + 10) = "Ship" Then Passed = False
             If MyGrid(Index + 20) = "Ship" Then Passed = False
      
          Case "Cruiser" 'Cruiser
             If MyGrid(Index - 10) = "Ship" Then Passed = False
             If MyGrid(Index) = "Ship" Then Passed = False
             If MyGrid(Index + 10) = "Ship" Then Passed = False
    
          Case "Submarine" 'Submarine
             If MyGrid(Index - 10) = "Ship" Then Passed = False
             If MyGrid(Index) = "Ship" Then Passed = False
             If MyGrid(Index + 10) = "Ship" Then Passed = False
    
          Case "Destroyer" 'Destroyer
             If MyGrid(Index - 10) = "Ship" Then Passed = False
             If MyGrid(Index) = "Ship" Then Passed = False
          
      End Select
    
    Else
       Select Case Ship 'Determine which ship to check
          Case "Carrier" 'Carrier ship
             If MyGrid(Index - 2) = "Ship" Then Passed = False
             If MyGrid(Index - 1) = "Ship" Then Passed = False
             If MyGrid(Index) = "Ship" Then Passed = False
             If MyGrid(Index + 1) = "Ship" Then Passed = False
             If MyGrid(Index + 2) = "Ship" Then Passed = False

          Case "Battleship" 'Battleship
             If MyGrid(Index - 1) = "Ship" Then Passed = False
             If MyGrid(Index) = "Ship" Then Passed = False
             If MyGrid(Index + 1) = "Ship" Then Passed = False
             If MyGrid(Index + 2) = "Ship" Then Passed = False
    
          Case "Cruiser" 'Cruiser ship
             If MyGrid(Index - 1) = "Ship" Then Passed = False
             If MyGrid(Index) = "Ship" Then Passed = False
             If MyGrid(Index + 1) = "Ship" Then Passed = False
 
          Case "Submarine" 'Submarine ship
             If MyGrid(Index - 1) = "Ship" Then Passed = False
             If MyGrid(Index) = "Ship" Then Passed = False
             If MyGrid(Index + 1) = "Ship" Then Passed = False

          Case "Destroyer" 'Destroyer ship
             If MyGrid(Index - 1) = "Ship" Then Passed = False
             If MyGrid(Index) = "Ship" Then Passed = False
          
       End Select
    End If
    
    'Do these checks for boundaries of grid

    'Don't have to do N S check just W E
    If Screen.MousePointer = ptrWE Then 'W E pointer
       
       'Divide and get remainder of position
       Border = Index Mod 10
       
       If Ship = "Carrier" Then 'Carrier ship
          If Border = 0 Then 'Ship was set improperly on grid
            Passed = False 'Change flag
          ElseIf Border = 1 Then 'Ship was set improperly on grid
            Passed = False 'Change flag
          ElseIf Border = 8 Then 'Ship was set improperly on grid
            Passed = False 'Change flag
          ElseIf Border = 9 Then 'Ship was set improperly on grid
            Passed = False 'Change flag
          End If
   
       ElseIf Ship = "Battleship" Then 'Battleship
          If Border = 0 Then 'Ship was set improperly on grid
            Passed = False 'Change flag
          ElseIf Border = 8 Then 'Ship was set improperly on grid
            Passed = False 'Change flag
          ElseIf Border = 9 Then 'Ship was set improperly on grid
            Passed = False 'Change flag
          End If

       ElseIf Ship = "Cruiser" Then 'Cruiser ship
          If Border = 0 Then 'Ship was set improperly on grid
            Passed = False 'Change flag
          ElseIf Border = 9 Then 'Ship was set improperly on grid
            Passed = False 'Change flag
          End If
    
       ElseIf Ship = "Submarine" Then 'Submarine
          If Border = 0 Then 'Ship was set improperly on grid
            Passed = False 'Change flag
          ElseIf Border = 9 Then 'Ship was set improperly on grid
            Passed = False 'Change flag
          End If
       ElseIf Ship = "Destroyer" Then 'Destroyer ship
          'Ship was set improperly on grid
          If Index Mod 10 = 0 Then
             Passed = False 'Change flag
          End If
       End If
    End If

    'Return value of function according to flag
    CheckPlacement = Passed
End Function

Sub cmdHelp_Click ()
    Dim Msg As String 'For message
    Dim NL As String 'For new line code

    NL = Chr$(10) + Chr$(13) 'Store new line code

    'Make message
    Msg = "Object of game is to sink all the computer's ships on the grid."
    Msg = Msg + NL + "To start a game select a skill from the list of skills"
    Msg = Msg + NL + "then click on the Start button."
    Msg = Msg + NL + ""
    Msg = Msg + NL + "You will be instructed to place down your ships,"
    Msg = Msg + NL + "Move the mouse over the place you want to place the"
    Msg = Msg + NL + "ship and click on the left mouse button."
    Msg = Msg + NL + "You can change the direction of a ship with a right"
    Msg = Msg + NL + "click of the mouse button."
    Msg = Msg + NL + ""
    Msg = Msg + NL + "Then you will be asked five question, for each one you"
    Msg = Msg + NL + "answer correctly you will be given the chance to shoot"
    Msg = Msg + NL + "at the computer.  Answer any wrong and the computer "
    Msg = Msg + NL + "gains a chance to shoot at you."
    Msg = Msg + NL + ""
    Msg = Msg + NL + "After you have answered your five questions click on"
    Msg = Msg + NL + "the grid to fire at the computers' ships."

    'Show message
    MsgBox Msg, 64, "OVERVIEW OF GAME"

    'Tell about program
    Msg = "Program written by Rodney Beede."
    Msg = Msg + NL + "Published by Infinisoft."
    Msg = Msg + NL + "E-mail me at rodney_beede@hotmail.com"
    Msg = Msg + NL + "Read the BatlMath.txt file for more information."

    MsgBox Msg, 64, "ABOUT GAME"
End Sub

Sub cmdResign_Click ()
    End 'End program
End Sub

Sub cmdStart_Click ()
    Dim X As Integer 'Used in counter

    tmrFlashGrid.Enabled = False 'Disable flashing grid

    'Fill in array for grid
    For X = 0 To 99
       MyGrid(X) = "Grid"
       PCGrid(X) = "Grid"
       MyShots(X) = "Grid"
       PCShots(X) = "Grid"
    Next X

    Call DrawGrid 'Call sub to draw empty grid

    lstSkills.Enabled = False 'Disable list box of skills
    lblSkill.Caption = "Skill is:" 'Change to know user what skill is

    lblInstructions.Caption = "Place your ships" 'Tell user what to do
    lblInstructions.Visible = True 'Show instuctions
    
    Started = True 'Set flag

    Ship = "Carrier" 'Set for making ships

    'Change mouse pointer appearance
    Call imgOceanGrid_MouseUp(0, 2, 0, 0, 0)

    Skill = lstSkills.ListIndex 'Set difficulty level

    cmdStart.Enabled = False 'Disable start button

End Sub

Sub DrawGrid ()
    Dim X As Integer 'Declare variable for counter
    
    'Fill in gird
    For X = 0 To 99
       If MyGrid(X) = "Ship" Then 'Place a ship here
          imgOceanGrid(X).Picture = imgShipBox
       ElseIf MyGrid(X) = "Hit" Then 'Place a hit box here
          imgOceanGrid(X).Picture = imbShiphBox
       ElseIf MyGrid(X) = "Miss" Then 'Place miss box here
          imgOceanGrid(X).Picture = imgMissBox
       ElseIf MyGrid(X) = "Grid" Then 'Place grid box here
          imgOceanGrid(X).Picture = imgGridBox
       End If
    Next X

End Sub

Sub Fire (Where As Integer)
    Dim X As Integer 'Declare variable for counter

    'Check to see if already shot their
    If MyShots(Where) = "Hit" Or MyShots(Where) = "Miss" Then
       'Tell user can't
       MsgBox "You can't shoot their again!", 16, "OH NO!"
       Exit Sub 'Leave sub
    End If
    
    If Not MyNumberShots = 0 Then 'Still have shots left
    
       'Fill in grid with place shot at
       For X = 0 To 99 'Start counter
          If MyShots(X) = "Hit" Then 'Put hit box here
             imgOceanGrid(X).Picture = imgHitBox
          ElseIf MyShots(X) = "Miss" Then 'Put miss box here
             imgOceanGrid(X).Picture = imgMissBox
          ElseIf MyShots(X) = "Grid" Then 'Put grid box here
             imgOceanGrid(X).Picture = imgGridBox
          End If
       Next X
       
       'Change label of grid
       lblGrid.Caption = "Place's you've shot at"
       
       If PCGrid(Where) = "Ship" Then 'Made a hit
          MyShots(Where) = "Hit" 'Store that hit was made
          MyHits = MyHits + 1 'Add hit to number of hits
          'Change picture to hit box on grid
          imgOceanGrid(Where).Picture = imgHitBox
          'Subtract number of shots allowed by one
          MyNumberShots = MyNumberShots - 1
          'Update label with number of shots left
          lblMyShots.Caption = "Shots you have left: " & MyNumberShots
          MsgBox "Hit", 48, "Got me" 'Tell player it was a hit
       ElseIf PCGrid(Where) = "Grid" Then 'Made a miss
          MyShots(Where) = "Miss" 'Store that miss was made
          imgOceanGrid(Where).Picture = imgMissBox 'Update grid
          'Subtract number of shots left by one
          MyNumberShots = MyNumberShots - 1
          'Update label with number of shots left
          lblMyShots.Caption = "Shots you have left: " & MyNumberShots
          'Tell user it was a miss
          MsgBox "Miss", 48, "Didn't get me"
       End If
    End If

    If Not PCNumberShots = 0 Then 'Computer has shots left
       
       'Fill grid with players ships
       For X = 0 To 99 'Start counter
          If MyGrid(X) = "Ship" Then 'Draw ship
             imgOceanGrid(X).Picture = imgShipBox
          ElseIf MyGrid(X) = "Hit" Then 'Draw a hit ship
             imgOceanGrid(X).Picture = imgShiphBox
          ElseIf MyGrid(X) = "Grid" Then 'Draw ocean
             imgOceanGrid(X).Picture = imgGridBox
          End If
       Next X
    
       'Change label for grid
       lblGrid.Caption = "Your Grid"
    
TryAgain:
       Where = (99 * Rnd) 'Get a random number
       
       'Check to see if alread shot their
       If PCShots(Where) = "Hit" Or PCShots(Where) = "Miss" Then
          GoTo TryAgain 'Did, go pick a different number
       End If
       
       If MyGrid(Where) = "Ship" Then 'PC made a hit
          PCShots(Where) = "Hit" 'Store hit in array where shot at
          PCHits = PCHits + 1 'Add to number of hits made
          MyGrid(Where) = "Hit" 'Store hit in player's array
          'Draw ship hit box on grid
          imgOceanGrid(Where).Picture = imgShiphBox
          'Decrease number of shots left
          PCNumberShots = PCNumberShots - 1
          'Update label with number of shots left
          lblPCShots.Caption = "Shots I have left: " & PCNumberShots
          'Tell player that a ship was hit
          MsgBox "You've been hit", 48, "I Got Ya"
       ElseIf MyGrid(Where) = "Grid" Then 'PC missed
          PCShots(Where) = "Miss" 'Store miss in array
          'Decrease number of shots left
          PCNumberShots = PCNumberShots - 1
          'Update label with number of shots left
          lblPCShots.Caption = "Shots I have left: " & PCNumberShots
          'Tell user PC missed
          MsgBox "I missed", 48, "I Didn't Get Ya"
       End If
    End If
    
    If MyHits = 17 Then 'Player has won
       'Tell player that he or she has won
       MsgBox "You have sunk all of the computer's ships!", 48, "YOU WIN, YOU WIN"
       End 'End program
    ElseIf PCHits = 17 Then 'Player has lost
       'Tell player that he or she has lost
       MsgBox "I have sunk all of your ships!", 48, "YOU LOSE, YOU LOSE"
       End 'End program
    End If

    If MyNumberShots + PCNumberShots = 0 Then 'Need to ask more questions
       Load frmQuestions 'Load up question form
       Exit Sub 'Leave sub
    End If
    
    If MyNumberShots = 0 Then 'Out of shots
       
       'Redraw grid for ships on grid
       For X = 0 To 99 'Start counter
          If MyGrid(X) = "Ship" Then 'Draw ship
             imgOceanGrid(X).Picture = imgShipBox
          ElseIf MyGrid(X) = "Hit" Then 'Draw hit ship
             imgOceanGrid(X).Picture = imgShiphBox
          ElseIf MyGrid(X) = "Grid" Then 'Draw ocean
             imgOceanGrid(X).Picture = imgGridBox
          End If
       Next X
    
       lblGrid.Caption = "Your Grid" 'Update label for grid
       'Give instructions
       lblInstructions.Caption = "Click on the grid"
    Else
       
       'Redraw grid for places shot at
       For X = 0 To 99 'Start counter
          If MyShots(X) = "Hit" Then 'Draw hit box
             imgOceanGrid(X).Picture = imgHitBox
          ElseIf MyShots(X) = "Miss" Then 'Draw miss box
             imgOceanGrid(X).Picture = imgMissBox
          ElseIf MyShots(X) = "Grid" Then 'Draw ocean
             imgOceanGrid(X).Picture = imgGridBox
          End If
       Next X
       
       'Tell user instructions
       lblInstructions.Caption = "Click on the grid to fire"
       'Change grid label
       lblGrid.Caption = "Place's you've shot at"
    End If

End Sub

Sub Form_Load ()
 
    'Set up position and size of form
    Me.Top = 0
    Me.Left = 0
    Me.Height = Screen.Height
    Me.Width = Screen.Width

    UpdatedAmt = 0 'Clear out variable
    Call Setup 'Call's a sub to setup the grid

    'Add skill levels to list of skills
    lstSkills.AddItem "2 + 2 = 5"
    lstSkills.AddItem "Not too hard"
    lstSkills.AddItem "I'm a genius"

End Sub

Sub imgOceanGrid_MouseMove (Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim C As Integer 'Used for counter
   Static oldIndex As Integer 'Stores old index
   
   On Error Resume Next 'Skip errors
   
   'Game has started still placing ships
   If Started = True And Not Ship = "" Then
      'Check to see if an update is needed
      If oldIndex = 0 Or oldIndex <> Index Or DirectionChange = True Then
          Call DrawGrid 'Call sub to draw current grid
          
          Select Case Ship 'Determine which ships to show
             Case "Carrier"  'Carrier ship
                If Screen.MousePointer = ptrWE Then 'Put ship going horizontal
                   imgOceanGrid(Index - 2).Picture = imgShipBox
                   imgOceanGrid(Index - 1).Picture = imgShipBox
                   imgOceanGrid(Index).Picture = imgShipBox
                   imgOceanGrid(Index + 1).Picture = imgShipBox
                   imgOceanGrid(Index + 2).Picture = imgShipBox
                ElseIf Screen.MousePointer = ptrNS Then 'Put ship going horizontal
                   imgOceanGrid(Index - 20).Picture = imgShipBox
                   imgOceanGrid(Index - 10).Picture = imgShipBox
                   imgOceanGrid(Index).Picture = imgShipBox
                   imgOceanGrid(Index + 10).Picture = imgShipBox
                   imgOceanGrid(Index + 20).Picture = imgShipBox
                End If
         
             Case "Battleship" 'Battleship
                If Screen.MousePointer = 9 Then 'Put ship going horizontal
                   imgOceanGrid(Index - 1).Picture = imgShipBox
                   imgOceanGrid(Index).Picture = imgShipBox
                   imgOceanGrid(Index + 1).Picture = imgShipBox
                   imgOceanGrid(Index + 2).Picture = imgShipBox
                ElseIf Screen.MousePointer = 7 Then 'Put ship going horizontal
                   imgOceanGrid(Index - 10).Picture = imgShipBox
                   imgOceanGrid(Index).Picture = imgShipBox
                   imgOceanGrid(Index + 10).Picture = imgShipBox
                   imgOceanGrid(Index + 20).Picture = imgShipBox
                End If
         
             Case "Cruiser" 'Cruiser ship
                If Screen.MousePointer = 9 Then 'Put ship going horizontal
                   imgOceanGrid(Index - 1).Picture = imgShipBox
                   imgOceanGrid(Index).Picture = imgShipBox
                   imgOceanGrid(Index + 1).Picture = imgShipBox
                ElseIf Screen.MousePointer = 7 Then 'Put ship going horizontal
                   imgOceanGrid(Index - 10).Picture = imgShipBox
                   imgOceanGrid(Index).Picture = imgShipBox
                   imgOceanGrid(Index + 10).Picture = imgShipBox
                End If
         
             Case "Submarine" 'Submarine
                If Screen.MousePointer = 9 Then 'Put ship going horizontal
                   imgOceanGrid(Index - 1).Picture = imgShipBox
                   imgOceanGrid(Index).Picture = imgShipBox
                   imgOceanGrid(Index + 1).Picture = imgShipBox
                ElseIf Screen.MousePointer = 7 Then 'Put ship going horizontal
                   imgOceanGrid(Index - 10).Picture = imgShipBox
                   imgOceanGrid(Index).Picture = imgShipBox
                   imgOceanGrid(Index + 10).Picture = imgShipBox
                End If
         
             Case "Destroyer" 'Destroyer
                If Screen.MousePointer = 9 Then 'Put ship going horizontal
                   imgOceanGrid(Index - 1).Picture = imgShipBox
                   imgOceanGrid(Index).Picture = imgShipBox
                ElseIf Screen.MousePointer = 7 Then 'Put ship going horizontal
                   imgOceanGrid(Index - 10).Picture = imgShipBox
                   imgOceanGrid(Index).Picture = imgShipBox
                End If
          End Select
      End If
   End If
   
   oldIndex = Index 'Store cell mouse was over
   
   'Check to see if done placing ships
   If Started = True And Shooting = False And Ship = "" Then
      Screen.MousePointer = ptrDefault 'Restore mouse pointer
      Call PCShipPlacement 'Call sub to put pc ships down
   End If
End Sub

Sub imgOceanGrid_MouseUp (Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Counter As Integer 'Declare variable for counter

    If Started = True And Not Ship = "" Then 'Game has started
 
       If Button = 2 Then
          'Right button was pressed. Change mouse icon
          If Screen.MousePointer = 7 Then
             Screen.MousePointer = 9 'W E pointer
          Else
             Screen.MousePointer = 7 'N S pointer
          End If
       
          DirectionChange = True 'Change flat
          
          'Force update on grid
          Call imgOceanGrid_MouseMove(Index, Button, Shift, X, Y)
    
       ElseIf Button = 1 Then
          'Left button was pressed. User wants to place ship on grid
          
          'See if ship's placement is valid
          If CheckPlacement(Index) = True Then
      
             'Ship's placement is valid
             'Store ship's position in array
         
             If Screen.MousePointer = ptrNS Then 'N S pointer
                Select Case Ship 'Determine which ship
                   Case "Carrier"  'Carrier ship
                      'Set up information on what to store in array
                      UpdatedAmt = 5
                      Updated(1) = Index - 20
                      Updated(2) = Index - 10
                      Updated(3) = Index
                      Updated(4) = Index + 10
                      Updated(5) = Index + 20
                      Ship = "Battleship" 'Change ship
             
                   Case "Battleship" 'Battleship
                      'Set up information on what to store in array
                      UpdatedAmt = 4
                      Updated(1) = Index - 10
                      Updated(2) = Index
                      Updated(3) = Index + 10
                      Updated(4) = Index + 20
                      Ship = "Cruiser" 'Change ship
            
                   Case "Cruiser" 'Cruiser ship
                   'Set up information on what to store in array
                      UpdatedAmt = 3
                      Updated(1) = Index - 10
                      Updated(2) = Index
                      Updated(3) = Index + 10
                      Ship = "Submarine" 'Change ship
            
                   Case "Submarine" 'Submarine
                      'Set up information on what to store in array
                      UpdatedAmt = 3
                      Updated(1) = Index - 10
                      Updated(2) = Index
                      Updated(3) = Index + 10
                      Ship = "Destroyer" 'Change ship
            
                   Case "Destroyer" 'Destroyer ship
                      'Set up information on what to store in array
                      UpdatedAmt = 2
                      Updated(1) = Index - 10
                      Updated(2) = Index
                      Ship = "" 'Change ship
                End Select
          
             ElseIf Screen.MousePointer = ptrWE Then
                'Ship is to be placed horizontaly
              
                Select Case Ship 'Determine which ship
                   Case "Carrier"  'Carrier ship
                      'Set up information on what to store in array
                      UpdatedAmt = 5
                      Updated(1) = Index - 2
                      Updated(2) = Index - 1
                      Updated(3) = Index
                      Updated(4) = Index + 1
                      Updated(5) = Index + 2
                      Ship = "Battleship" 'Change ship
            
                   Case "Battleship" 'Battleship
                      'Set up information on what to store in array
                      UpdatedAmt = 4
                      Updated(1) = Index - 1
                      Updated(2) = Index
                      Updated(3) = Index + 1
                      Updated(4) = Index + 2
                      Ship = "Cruiser" 'Change ship
                
                   Case "Cruiser" 'Cruiser ship
                      'Set up information on what to store in array
                      UpdatedAmt = 3
                      Updated(1) = Index - 1
                      Updated(2) = Index
                      Updated(3) = Index + 1
                      Ship = "Submarine" 'Change ship
                
                   Case "Submarine" 'Submarine
                      'Set up information on what to store in array
                      UpdatedAmt = 3
                      Updated(1) = Index - 1
                      Updated(2) = Index
                      Updated(3) = Index + 1
                      Ship = "Destroyer" 'Change ship
                
                   Case "Destroyer" 'Destroyer ship
                      'Set up information on what to store in array
                      UpdatedAmt = 2
                      Updated(1) = Index - 1
                      Updated(2) = Index
                      Ship = "" 'Change ship
              
                End Select
             End If
             
             'Start a counter to array
             For Counter = 1 To UpdatedAmt
                MyGrid(Updated(Counter)) = "Ship"
             Next Counter
             
             'Redraw part of grid with new ship on it
             Call UpdateGrid
          Else
             'Invalid placement, tell user
             MsgBox "You may not place a ship there!", 48, "Error"
          End If
       End If
    End If

    If Shooting = True Then 'Player is shooting on grid
       Call Fire(Index) 'Call sub to fire
    End If
End Sub

Sub lblInstructions_Change ()
    'Center instructions on screen
    lblInstructions.Left = (Me.Width / 2) - (lblInstructions.Width / 2)
    
    If lblInstructions.Caption = "" Then 'No instructions to show
       lblInstructions.Visible = False 'Hide label
    Else
       lblInstructions.Visible = True 'Show label
    End If
End Sub

Sub lblInstructions_Click ()
    Dim Counter As Integer 'Declare for counter
    
    If lblInstructions.Caption = "Click here to begin" Then
       'Get grid ready for player to fire
       
       For Counter = 0 To 99 'Start counter
          If MyShots(Counter) = "Hit" Then 'Change picture in box
             imgOceanGrid(Counter).Picture = imgHitBox
          ElseIf MyShots(Counter) = "Miss" Then 'Change picture in box
             imgOceanGrid(Counter).Picture = imgMissBox
          ElseIf MyShots(Counter) = "Grid" Then 'Change picture in box
             imgOceanGrid(Counter).Picture = imgGridBox
          End If
       Next Counter
       
       If MyNumberShots = 0 Then 'User is out of shots
          'Change labels to tell user what to do
          lblInstructions.Caption = "Click on the grid"
          lblGrid.Caption = "Your Grid"
       Else
          'User still has shots
          'Tell user what to do
          lblInstructions.Caption = "Click on the grid to fire."
          'Update label for grid
          lblGrid.Caption = "Places you've shot at"
       End If
       
       lblGrid.Visible = True 'Show grid label
       Shooting = True 'Set flag
    End If
End Sub

Sub lstSkills_Click ()
    cmdStart.Enabled = True 'Enable start button
End Sub

Sub PCShipPlacement ()
    Dim Direction As Integer 'Declare variable
    Dim Position As Integer 'Declare variable
    Dim Counter As Integer 'Declare variable
    
    Direction = Int(2 * Rnd) 'Choose random number and store it
    Position = Int(5 * Rnd + 1) 'Choose random number

    'Store Carrier position on grid for PC
    If Direction = 0 Then 'Place ship going horizontal
       If Position = 1 Then 'Place ship on 6th row
          For Counter = 55 To 59 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 2 Then 'Place ship on 7th row
          For Counter = 65 To 69 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 3 Then 'Place ship on 8th row
          For Counter = 75 To 79 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 4 Then 'Place ship on 9th row
          For Counter = 85 To 89 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 5 Then 'Place ship on 10th row
          For Counter = 95 To 99 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       End If
    ElseIf Direction = 1 Then 'Place ship going vertical
       If Position = 1 Then 'Place ship on 6th column
          For Counter = 55 To 95 Step 10 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 2 Then 'Place ship on 7th column
          For Counter = 56 To 96 Step 10 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 3 Then 'Place ship on 8th column
          For Counter = 57 To 97 Step 10 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 4 Then 'Place ship on 9th column
          For Counter = 58 To 98 Step 10 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 5 Then 'Place ship on 10th column
          For Counter = 59 To 99 Step 10 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       End If
    End If

    Direction = Int(2 * Rnd) 'Choose random number and store it
    Position = Int(4 * Rnd + 1) 'Choose random number

    'Store Battleship position on grid for PC
    If Direction = 0 Then 'Place ship going horizontal
       If Position = 1 Then 'Place ship on 7th row
          For Counter = 60 To 63 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 2 Then 'Place ship on 8th row
          For Counter = 70 To 73 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 3 Then 'Place ship on 9th row
          For Counter = 80 To 83 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 4 Then 'Place ship on 10th row
          For Counter = 90 To 93 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       End If
    ElseIf Direction = 1 Then 'Place ship going vertical
       If Position = 1 Then 'Place ship on 7th column
          For Counter = 60 To 90 Step 10 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 2 Then 'Place ship on 8th column
          For Counter = 61 To 91 Step 10 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 3 Then 'Place ship on 9th column
          For Counter = 62 To 92 Step 10 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 4 Then 'Place ship on 10th column
          For Counter = 63 To 93 Step 10 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       End If
    End If

    Direction = Int(2 * Rnd) 'Choose random number and store it
    Position = Int(3 * Rnd + 1) 'Choose random number

    'Store Cruiser position on grid for PC
    If Direction = 0 Then 'Place ship going horizontal
       If Position = 1 Then 'Place ship on 1st row
          For Counter = 0 To 2 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 2 Then 'Place ship on 2nd row
          For Counter = 10 To 12 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 3 Then 'Place ship on 3rd row
          For Counter = 20 To 22 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       End If
    ElseIf Direction = 1 Then 'Place ship going vertical
       If Position = 1 Then 'Place ship on 1st column
          For Counter = 0 To 20 Step 10 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 2 Then 'Place ship on 2nd column
          For Counter = 1 To 21 Step 10 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 3 Then 'Place ship on 3rd column
          For Counter = 2 To 31 Step 10 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       End If
    End If

    Direction = Int(2 * Rnd) 'Choose random number and store it
    Position = Int(3 * Rnd + 1) 'Choose random number

    'Store Submarine position on grid for PC
    If Direction = 0 Then 'Place ship going horizontal
       If Position = 1 Then 'Place ship on 1st row
          For Counter = 7 To 9 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 2 Then 'Place ship on 2nd row
          For Counter = 17 To 19 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 3 Then 'Place ship on 3rd row
          For Counter = 27 To 29 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       End If
    ElseIf Direction = 1 Then 'Place ship going vertical
       If Position = 1 Then 'Place ship on 8th column
          For Counter = 7 To 27 Step 10 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 2 Then 'Place ship on 9th column
          For Counter = 8 To 28 Step 10 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 3 Then 'Place ship on 10th column
          For Counter = 9 To 29 Step 10 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       End If
    End If

    Direction = Int(2 * Rnd) 'Choose random number and store it
    Position = Int(2 * Rnd + 1) 'Choose random number

    'Store Submarine position on grid for PC
    If Direction = 0 Then 'Place ship going horizontal
       If Position = 1 Then 'Place ship on 4th row
          For Counter = 33 To 34 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 2 Then 'Place ship on 5th row
          For Counter = 43 To 44 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       End If
    ElseIf Direction = 1 Then 'Place ship going vertical
       If Position = 1 Then 'Place ship on 4th column
          For Counter = 33 To 43 Step 10 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       ElseIf Position = 2 Then 'Place ship on 5th column
          For Counter = 34 To 44 Step 10 'Start counter
             PCGrid(Counter) = "Ship" 'Fill in grid
          Next Counter
       End If
    End If
    
    frmMain!lblInstructions.Caption = "" 'Erase instructions
    Load frmQuestions 'Load questions up
End Sub

Sub Setup ()
   
   Const iLeftSide = 2280 'Left side for grid
   Dim X As Integer 'Used in counter
   Dim iHeightOfCell As Integer 'Height of each block
  
   frmMain!imgOceanGrid(0).Left = iLeftSide  'Set starting grid position
   frmMain!imgOceanGrid(0).Top = 960         'Set starting grid position
   iHeightOfCell = imgOceanGrid(0).Height    ' height of all cells are the same
   
   For X = 1 To 99 'Start counter
      Load frmMain!imgOceanGrid(X) 'Load new cell
   
      'Checks to see if new line is starting
      'If so then resets new block to go below the last line
      If X Mod 10 = 0 Then
         frmMain!imgOceanGrid(X).Left = iLeftSide
         frmMain!imgOceanGrid(X).Top = frmMain!imgOceanGrid(X - 10).Top + iHeightOfCell
      Else 'Not doing a new line keep adding to current one
         frmMain!imgOceanGrid(X).Left = frmMain!imgOceanGrid(X - 1).Left + frmMain!imgOceanGrid(X - 1).Width
         frmMain!imgOceanGrid(X).Top = frmMain!imgOceanGrid(X - 1).Top
      End If
   
      frmMain!imgOceanGrid(X).Visible = True 'Show new block
   Next X

End Sub

Sub tmrFlashGrid_Timer ()

    Randomize 'Initializes the random-number generator

    'Draw hit, miss, and grid boxes at random spots
    imgOceanGrid(Int(99 * Rnd + 1)).Picture = imgHitBox
    imgOceanGrid(Int(99 * Rnd + 1)).Picture = imgHitBox
    imgOceanGrid(Int(99 * Rnd)).Picture = imgMissBox
    imgOceanGrid(Int(99 * Rnd)).Picture = imgMissBox
    imgOceanGrid(Int(99 * Rnd)).Picture = imgGridBox

End Sub

Sub UpdateGrid ()

   Dim X As Integer 'For counter
   Dim CellToUpdate As Integer 'For storing which cell to update
   
   'Update grid
   For X = 1 To UpdatedAmt
      CellToUpdate = Updated(X) 'Current cell to change
      'Change picture in grid
      imgOceanGrid(CellToUpdate).Picture = imgShipBox
   Next X

End Sub

