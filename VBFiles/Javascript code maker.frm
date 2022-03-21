VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Javascript"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3390
   Icon            =   "Javascript code maker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   3390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.ListBox lstFiles 
      Height          =   1620
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton cmdDateTime 
      Caption         =   "Date && Time Display"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   4320
      Width           =   2895
   End
   Begin VB.CommandButton cmdBackGnd 
      Caption         =   "Random Background Picture"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton cmdBackSnd 
      Caption         =   "Random Background Sound"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label lblFiles 
      AutoSize        =   -1  'True
      Caption         =   "Files:"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   360
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      Caption         =   "Web page path to files:"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   1650
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    If txtFileName.Text = "" Then Exit Sub  'Don't add anything
    
    'Make sure their are not any special ' characters in the filename
    If Not InStr(txtFileName.Text, "'") = False Then
        'Tell user invalid name
        MsgBox "Invalid filename.  You cannot have any ' characters in the filename.", vbExclamation, "Error"
        Exit Sub  'Leave sub
    End If
    
    lstFiles.AddItem txtFileName.Text  'Add filename to list
    txtFileName.Text = ""  'Clear out text
    
    cmdDelete.Enabled = True  'Enabled delete button
End Sub

Private Sub cmdBackGnd_Click()
    Dim strCode As String  'For code
    Dim i As Integer  'For counter
    Dim Pass As Integer  'For a flag
    
    'Figure what this button is doing
    If cmdBackGnd.Caption = "Random Background Picture" Then
        cmdBackGnd.Caption = "Code It!"  'Reset caption
        
        'Give user instructions
        MsgBox "Fill in the path of where the graphic files will be on the web server (blank means same place as document, also make sure to put a \ at the end), then type in the name of the graphic files to be displayed.  When you are done click on the Code It! button.", vbInformation, "Javascript"
    
        'Clear out boxes and list
        txtPath.Text = ""
        txtFileName.Text = ""
        lstFiles.Clear
        
        'Enabled boxes
        txtPath.Enabled = True
        txtFileName.Enabled = True
        lstFiles.Enabled = True
        cmdAdd.Enabled = True
        
        'Disable other buttons
        cmdBackSnd.Enabled = False
        cmdDateTime.Enabled = False
    
        'Color in boxes to look enabled
        txtPath.BackColor = vbWhite
        txtFileName.BackColor = vbWhite
        lstFiles.BackColor = vbWhite
        
    Else  'Coding
        cmdBackGnd.Caption = "Random Background Picture"  'Reset caption
    
        Pass = False  'Assume not done
        
        'Format path information into correct format
        Do While Pass = False
            If txtPath.Text = "" Then Exit Do  'Leave loop
            
            For i = 1 To Len(txtPath.Text)
                If Mid$(txtPath.Text, i, 1) = "\" Then
                    'Check for another slash
                    If Mid$(txtPath.Text, i + 1, 1) = "\" Then
                        'Assume done
                        Pass = True
                        i = i + 1  'Move on
                    Else  'Need to add a slash
                        txtPath.Text = Left$(txtPath.Text, i) + "\" + Mid$(txtPath.Text, i + 1, Len(txtPath.Text))
                        Pass = False  'Not done
                        Exit For  'Leave loop
                    End If
                End If
            Next i
        Loop
    
        'Make the code
        strCode = "<!-- Stuff for Background -->" + vbCrLf
        strCode = strCode + vbCrLf
        strCode = strCode + "<SCRIPT language='javascript'>" + vbCrLf
        strCode = strCode + "<!-- hide" + vbCrLf
        strCode = strCode + "  var howMany =" + Str$(lstFiles.ListCount) + vbCrLf
        strCode = strCode + "  var pic = new Array(howMany)" + vbCrLf
        strCode = strCode + vbCrLf
    
        'Add all the file names
        For i = 1 To lstFiles.ListCount
            strCode = strCode + "  pic[" + Str$(i) + "]='" + lstFiles.List(i - 1) + "'" + vbCrLf
        Next i

        'Continue on with code
        strCode = strCode + "  function rndnumber(){" + vbCrLf
        strCode = strCode + "        var randscript = -1" + vbCrLf
        strCode = strCode + "        while (randscript <= 0 || randscript > howMany || isNaN(randscript)){" + vbCrLf
        strCode = strCode + "                randscript = parseInt(Math.random()*(howMany+1))" + vbCrLf
        strCode = strCode + "        }" + vbCrLf
        strCode = strCode + "        return randscript" + vbCrLf
        strCode = strCode + "  }" + vbCrLf
        strCode = strCode + vbCrLf
        strCode = strCode + "     var randomsub = rndnumber()" + vbCrLf
        strCode = strCode + "     var pfile = pic[randomsub]" + vbCrLf
        strCode = strCode + vbCrLf
        strCode = strCode + "  pfile = '" + txtPath.Text + "' + pfile" + vbCrLf
        strCode = strCode + "  document.write('<BODY BACKGROUND=' + pfile + '>')" + vbCrLf
        strCode = strCode + vbCrLf
        strCode = strCode + "//unhide -->" + vbCrLf
        strCode = strCode + "</SCRIPT>" + vbCrLf
        strCode = strCode + vbCrLf
        strCode = strCode + "<BODY>" + vbCrLf
    
        'Check to see if thier were any files
        If lstFiles.ListCount > 0 Then  'Their were, give user code
            'Copy the code
            Clipboard.Clear  'Clear out clipboard
            Clipboard.SetText strCode  'Send data to clipboard
            
            'Tell user what to do
            MsgBox "The code has been copied to the clipboard.  To use it paste the code right after the </HEAD> tag in your html document.", vbInformation, "Javascript"
        Else  'Nothing for user
            MsgBox "You didn't provide any file names, no code made.", vbExclamation, "Error"
        End If
        
        'Clear out boxes and list
        txtPath.Text = ""
        txtFileName.Text = ""
        lstFiles.Clear
        
        'Disable boxes
        txtPath.Enabled = False
        txtFileName.Enabled = False
        lstFiles.Enabled = False
        cmdAdd.Enabled = False
        cmdDelete.Enabled = False
        
        'Enable other buttons
        cmdBackSnd.Enabled = True
        cmdDateTime.Enabled = True
    
        'Color in boxes to look disabled
        txtPath.BackColor = frmMain.BackColor
        txtFileName.BackColor = frmMain.BackColor
        lstFiles.BackColor = frmMain.BackColor
    End If
End Sub

Private Sub cmdBackSnd_Click()
    Dim strCode As String  'For storing coding
    Dim i As Integer  'For counter
    Dim Pass As Integer  'For flag checking
    
    'Check to see if starting to make code
    If cmdBackSnd.Caption = "Random Background Sound" Then
        cmdBackSnd.Caption = "Code It!"  'Reset caption
    
        'Give user instructions
        MsgBox "Fill in the path of where the sound files will be on the web server (blank means same place as document, also make sure to put a \ at the end), then type in the name of the sound files to be displayed.  When you are done click on the Code It! button.", vbInformation, "Javascript"
    
        'Clear out boxes and list
        txtPath.Text = ""
        txtFileName.Text = ""
        lstFiles.Clear
        
        'Enabled boxes
        txtPath.Enabled = True
        txtFileName.Enabled = True
        lstFiles.Enabled = True
        cmdAdd.Enabled = True
        
        'Disable other buttons
        cmdBackGnd.Enabled = False
        cmdDateTime.Enabled = False
    
        'Color in boxes to look enabled
        txtPath.BackColor = vbWhite
        txtFileName.BackColor = vbWhite
        lstFiles.BackColor = vbWhite
    Else  'Not
        cmdBackSnd.Caption = "Random Background Sound"  'Reset caption
    
        Pass = False  'Assume not done
        
        'Format path information into correct format
        Do While Pass = False
            If txtPath.Text = "" Then Exit Do  'Nothing to modify
                        
            For i = 1 To Len(txtPath.Text)
                If Mid$(txtPath.Text, i, 1) = "\" Then
                    'Check for another slash
                    If Mid$(txtPath.Text, i + 1, 1) = "\" Then
                        'Assume done
                        Pass = True
                        i = i + 1  'Move on
                    Else  'Need to add a slash
                        txtPath.Text = Left$(txtPath.Text, i) + "\" + Mid$(txtPath.Text, i + 1, Len(txtPath.Text))
                        Pass = False  'Not done
                        Exit For  'Leave loop
                    End If
                End If
            Next i
        Loop

        'Make up the javascript code
        strCode = "<! Random Background sound code starts below... ><SCRIPT language='javascript'><!-- hide" + vbCrLf
        strCode = strCode + "var howMany = " + Str$(lstFiles.ListCount) + vbCrLf
        strCode = strCode + "var snd = new Array(howMany)" + vbCrLf
        strCode = strCode + vbCrLf
        
        'Add the list of files
        For i = 1 To lstFiles.ListCount
            strCode = strCode + "snd[" + Str$(i) + "]='" + lstFiles.List(i - 1) + "'" + vbCrLf
        Next i
        
        'Continue on with the rest of the code
        strCode = strCode + vbCrLf
        strCode = strCode + "function rndnumber(){" + vbCrLf
        strCode = strCode + "        var randscript = -1" + vbCrLf
        strCode = strCode + "        while (randscript <= 0 || randscript > howMany || isNaN(randscript)){" + vbCrLf
        strCode = strCode + "                randscript = parseInt(Math.random()*(howMany+1))" + vbCrLf
        strCode = strCode + "        }" + vbCrLf
        strCode = strCode + "        return randscript" + vbCrLf
        strCode = strCode + "}" + vbCrLf
        strCode = strCode + vbCrLf
        strCode = strCode + "     var randomsub = rndnumber()" + vbCrLf
        strCode = strCode + "     var sfile = snd[randomsub]" + vbCrLf
        strCode = strCode + vbCrLf
        
        strCode = strCode + "    sfile= '" + txtPath.Text + "' + sfile" + vbCrLf
        strCode = strCode + vbCrLf
        strCode = strCode + "    document.write ('<bgsound src=' + sfile + '>')" + vbCrLf
        strCode = strCode + vbCrLf
        strCode = strCode + "    if (navigator.appName == 'Netscape') {" + vbCrLf
        strCode = strCode + "        document.write ('<EMBED src= ' + sfile + ' hidden=True autostart=true>')" + vbCrLf
        strCode = strCode + "    }" + vbCrLf
        strCode = strCode + vbCrLf
        strCode = strCode + "    document.write ('<BR>Out of ' + howMany + ' songs you are hearing song number ' + randomsub + '.<BR>The name of this song is ' + snd[randomsub] + '.')" + vbCrLf
        strCode = strCode + " --></SCRIPT><!...end of Random Background Sound code></P>" + vbCrLf
        strCode = strCode + vbCrLf
        strCode = strCode + "<! Rest of web page begins here -->"
    
        'Check to see if their are any files
        If lstFiles.ListCount > 0 Then  'There are
            'Put text in clipboard
            Clipboard.Clear  'Clear clipboard
            Clipboard.SetText strCode  'Put code in
            
            'Tell user what to do
            MsgBox "The code has been copied to the clipboard, to use it paste it right before your </BODY> closing tag in your file.", vbInformation, "Javascript"
        Else  'None, tell user
            MsgBox "You didn't provide any files, no code made.", vbExclamation, "Error"
        End If
        'Clear out boxes and list
        txtPath.Text = ""
        txtFileName.Text = ""
        lstFiles.Clear
        
        'Disable boxes
        txtPath.Enabled = False
        txtFileName.Enabled = False
        lstFiles.Enabled = False
        cmdAdd.Enabled = False
        cmdDelete.Enabled = False
        
        'Enable other buttons
        cmdBackGnd.Enabled = True
        cmdDateTime.Enabled = True
    
        'Color in boxes to look disabled
        txtPath.BackColor = frmMain.BackColor
        txtFileName.BackColor = frmMain.BackColor
        lstFiles.BackColor = frmMain.BackColor
    End If
End Sub

Private Sub cmdDateTime_Click()
    Dim strCode As String  'For storing code
    
    strCode = "<! First Section >"
    strCode = strCode + "<SCRIPT language='javascript'>" + vbCrLf
    strCode = strCode + "<!---[JAVASCRIPT]---" + vbCrLf
    strCode = strCode + vbCrLf
    strCode = strCode + "var Clock = 0;" + vbCrLf
    strCode = strCode + vbCrLf
    strCode = strCode + "function Update() {" + vbCrLf
    strCode = strCode + "   var TimeObject = new Date();" + vbCrLf
    strCode = strCode + "   var TimeString = ' ';" + vbCrLf
    strCode = strCode + "   var DateString = '  ';" + vbCrLf
    strCode = strCode + "   var ClockHours = TimeObject.getHours();" + vbCrLf
    strCode = strCode + "   var ClockMinutes = TimeObject.getMinutes();" + vbCrLf
    strCode = strCode + "   var ClockSeconds = TimeObject.getSeconds();" + vbCrLf
    strCode = strCode + "   var ClockMonth = TimeObject.getMonth() + 1;" + vbCrLf
    strCode = strCode + "   var ClockDate = TimeObject.getDate();" + vbCrLf
    strCode = strCode + "   var ClockYear = TimeObject.getYear();" + vbCrLf
    strCode = strCode + "   var ClockAmPm;" + vbCrLf
    strCode = strCode + vbCrLf
    strCode = strCode + "   if( ClockHours < 12 ){" + vbCrLf
    strCode = strCode + "     if( !ClockHours ) {" + vbCrLf
    strCode = strCode + "       ClockHours = 12;" + vbCrLf
    strCode = strCode + "       ClockAmPm = 'M';" + vbCrLf
    strCode = strCode + "       }" + vbCrLf
    strCode = strCode + "     else {" + vbCrLf
    strCode = strCode + "       ClockAmPm = 'AM';" + vbCrLf
    strCode = strCode + "       }" + vbCrLf
    strCode = strCode + "     }" + vbCrLf
    strCode = strCode + "   else {" + vbCrLf
    strCode = strCode + "     if( ClockHours == 12 ) {" + vbCrLf
    strCode = strCode + "       ClockAmPm = 'PM';" + vbCrLf
    strCode = strCode + "       }" + vbCrLf
    strCode = strCode + "     else {" + vbCrLf
    strCode = strCode + "       ClockHours -= 12;" + vbCrLf
    strCode = strCode + "       ClockAmPm = 'PM';" + vbCrLf
    strCode = strCode + "       }" + vbCrLf
    strCode = strCode + "     }" + vbCrLf
    strCode = strCode + "   TimeString += ((ClockHours < 10) ? ' ' : '') + ClockHours + ':';" + vbCrLf
    strCode = strCode + "   TimeString += ((ClockMinutes < 10) ? '0' : '') + ClockMinutes + ':';" + vbCrLf
    strCode = strCode + "   TimeString += ((ClockSeconds < 10) ? '0' : '') + ClockSeconds + '';" + vbCrLf
    strCode = strCode + "   TimeString += ' ';" + vbCrLf
    strCode = strCode + "   TimeString += ClockAmPm + ' ';" + vbCrLf
    strCode = strCode + vbCrLf
    strCode = strCode + "   if( ClockMonth < 10 ) {" + vbCrLf
    strCode = strCode + "     ClockMonth = '0' + ClockMonth;" + vbCrLf
    strCode = strCode + "     }" + vbCrLf
    strCode = strCode + "   if( ClockDate < 10 ) {" + vbCrLf
    strCode = strCode + "     ClockDate = '0' + ClockDate;" + vbCrLf
    strCode = strCode + "     }" + vbCrLf
    strCode = strCode + "   DateString += ClockMonth+'/'+ClockDate+'/'+ClockYear;" + vbCrLf
    strCode = strCode + vbCrLf
    strCode = strCode + "   document.timeForm.dateBox.value = DateString;" + vbCrLf
    strCode = strCode + "   document.timeForm.timeBox.value = TimeString;" + vbCrLf
    strCode = strCode + vbCrLf
    strCode = strCode + "   clearTimeout( Clock );" + vbCrLf
    strCode = strCode + "   Clock = setTimeout( 'Update()', 1000 );" + vbCrLf
    strCode = strCode + "   }" + vbCrLf
    strCode = strCode + "//---[JAVASCRIPT]-->" + vbCrLf
    strCode = strCode + "</SCRIPT>" + vbCrLf
    strCode = strCode + vbCrLf
    strCode = strCode + vbCrLf
    strCode = strCode + "<! Second Section >" + vbCrLf
    strCode = strCode + "<!Add this section in the body>" + vbCrLf
    strCode = strCode + "<CENTER><FORM NAME='timeForm'>" + vbCrLf
    strCode = strCode + "<INPUT TYPE='text' NAME='dateBox' VALUE='' size='11'>" + vbCrLf
    strCode = strCode + "<INPUT TYPE='text' NAME='timeBox' VALUE='' size='13'>" + vbCrLf
    strCode = strCode + "</FORM></CENTER>" + vbCrLf
    
    Clipboard.Clear  'Clear clipboard
    Clipboard.SetText strCode   'Send code to clipboard
    
    'Tell user what to do now
    MsgBox "The code has been copied to the clipboard.  Copy the first section inbetween the </HEAD> tag and the <BODY> tag.  Cut the second section and paste it anywhere inside the body you wish.  Make your <BODY> tag look like this <BODY onLoad='Update()'>.", vbInformation, "Javascript"
End Sub

Private Sub cmdDelete_Click()
    If lstFiles.ListIndex = -1 Then Exit Sub  'Nothing selected
    
    lstFiles.RemoveItem lstFiles.ListIndex  'Remove item

    If lstFiles.ListCount = 0 Then cmdDelete.Enabled = False  'Out of things to delete
End Sub

Private Sub Form_Load()
    'Disable fill in sections for files
    txtPath.Enabled = False
    txtFileName.Enabled = False
    lstFiles.Enabled = False
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    
    'Color in boxes to look disabled
    txtPath.BackColor = frmMain.BackColor
    txtFileName.BackColor = frmMain.BackColor
    lstFiles.BackColor = frmMain.BackColor
    
    Me.Show  'Show form
    
    'Display what the user can do
    MsgBox "Select a function you would like to do for a javascript example.", vbInformation, "Javascript"
End Sub


Private Sub txtFileName_Change()
    Dim i As Integer  'For counter
    Dim Length As Integer  'For lentgh of string
    
    Length = Len(txtFileName.Text)  'Set length
    
    Do While Not Length <= 0
        'Loop through until a good match can be found in list
        For i = 0 To lstFiles.ListCount - 1
            If Left$(txtFileName.Text, Length) = Left$(lstFiles.List(i), Length) Then
                'Good match, select it
                lstFiles.ListIndex = i
                Length = -1  'Send flag to leave do-loop
                Exit For  'Leave for loop
            End If
        Next i
    
        Length = Length - 1  'One less in string
    Loop

    If txtFileName.Text = "" Then lstFiles.ListIndex = -1  'Select nothing in the list
End Sub
