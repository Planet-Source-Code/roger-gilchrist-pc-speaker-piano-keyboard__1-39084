VERSION 5.00
Begin VB.Form FrmKBDemo 
   Caption         =   "Key Board Class Demo"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Help"
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   10
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete Last"
      Height          =   375
      Index           =   3
      Left            =   9480
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox cmbTempo 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Texts"
      Height          =   375
      Index           =   2
      Left            =   8160
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox cmbLooks 
      Height          =   315
      Left            =   9360
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play Basica"
      Height          =   375
      Index           =   1
      Left            =   9480
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   9375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play Nokia"
      Height          =   375
      Index           =   0
      Left            =   9480
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9375
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   10515
      TabIndex        =   0
      Top             =   1320
      Width           =   10575
   End
   Begin VB.Label Label2 
      Caption         =   "Duration (Mouse Button) L=16 R=32   M=64        Shift+L=2  Shift+R=4  Shift+M=8      Ctrl+Any=1"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   2400
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "KeyBoard Style"
      Height          =   255
      Left            =   8280
      TabIndex        =   9
      Top             =   2280
      Width           =   1095
   End
End
Attribute VB_Name = "FrmKBDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright 2002 Roger Gilchrist
'email: rojagilkrist@hotmail.com

'This is just a quick and dirty demo of ClsKeyBoardPicture
'Thanks to Nokia Ringtone Player by Ovidiu Daniel Diaconescu for the original inspiration and the
'Data in the NoteArray and some of the code to make noise

Option Explicit
Private KB As New ClsKeyBoardPicture
Private Declare Sub InitCommonControls Lib "comctl32" () ':( Line inserted by Formatter

Private Sub cmbLooks_Click()
    KB.KeyBoardLook cmbLooks.ListIndex
End Sub

Private Sub cmbTempo_Click()
    KB.Tempo = cmbTempo.Text
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0 'Play Nokia
KB.ReadNokia Text1(0)
Case 1 'Play Basica
KB.ReadBasica Text1(1)
Case 2 'Clear Texts
Text1(0) = ""
Text1(1) = ""
Case 3 'Delete last
If Len(Text1(0)) Then ' is text
If InStr(Text1(0), " ") Then 'more than one note
    Text1(0).Text = Left(Text1(0).Text, InStrRev(Text1(0).Text, " ") - 1)
    Text1(1).Text = Left(Text1(1).Text, InStrRev(Text1(1).Text, " ") - 1)
Else 'only one note
    Text1(0) = ""
    Text1(1) = ""
End If
End If
Case 4
'Nokia format Pause <Duration>P Note = <Duration>[#]<Note><Octave>
'Basica       "Pause = P<Duration>  Note = L<Duration> O<Octave> [#]<Note>"

MsgBox "This is just a quicky Demo of PC Speaker music and my keyboard class." & vbCr & _
"The keyboard is stretchable so just drag form border to expand it." & vbCr & _
"WARNING: The sound is asynchronous (while a sound or pause is playing you cannot do anything else)" & vbCr & _
"Basica Format: Pause = P<Duration>  Note = L<Duration> O<Octave><Note>[#]  Flats= <PreviousNote>#" & vbCr & _
"                   This is a Simplified version Basica's Play command code." & vbCr & _
"Nokia Format:  Pause = <Duration>P   Note = <Duration>[#]<Note><Octave>  Flats= #<PreviousNote>.  " & vbCr & _
"                   This code may also be simplified. I only know the code from the Ovidiu's program." & vbCr & _
"You can edit the text manually but there is no error checking and the slightest error is a crash because it uses API." & vbCr & _
"If you build a tune you actually want to keep just cut and paste it and don't forget to add the tempo value." & vbCr & _
"If there is interest will develop further." & vbCr & _
 vbCr & _
"Copyright 2002 Roger Gilchrist email: rojagilkrist@hotmail.com" & vbCr & _
 vbCr & _
"Thanks to Ovidiu Daniel Diaconescu for inspiration and noise code"


End Select
End Sub

Private Sub Form_Initialize() ':) Line inserted by Formatter

    InitCommonControls ':) Line inserted by Formatter

End Sub ':) Line inserted by Formatter

Private Sub Form_Load()
Dim I As Integer
Me.Width = Screen.Width * 2 / 3
With cmbLooks
.AddItem "Antique"
.AddItem "Classical"
.AddItem "Default"
.Text = "Default"
End With
With cmbTempo
For I = 32 To 255
.AddItem I
Next
.Text = 120
End With
    With KB
        Set .AssignControl = Picture1
        .UseFormCaption = True
        .PauseKeyOn = True
    End With 'KB
   
End Sub

Private Sub Form_Resize()
' place the form in the any third of the screen
'and resize slightly to see different KeyBoardLooks
With FrmKBDemo
    Picture1.Width = FrmKBDemo.ScaleWidth
    Text1(0).Width = FrmKBDemo.ScaleWidth - Command1(0).Width
    Text1(1).Width = FrmKBDemo.ScaleWidth - Command1(1).Width
    Command1(0).Left = Text1(0).Width
    Command1(1).Left = Text1(1).Width
    Command1(2).Left = Text1(1).Width
    cmbLooks.Left = Text1(1).Width
    Label1.Left = cmbLooks.Left - Label1.Width
    Command1(3).Left = Text1(1).Width - Command1(3).Width
    
    End With 'PICTURE1
    KB.Resize

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    KB.MouseDown Button, Shift, X, Y

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    KB.MouseMove Button, Shift, X, Y

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    KB.MouseUp Button, Shift, X, Y
    KB.TextOutPutStyle = Nokia
If Len(KB.TextOutPut) Then
Text1(0).Text = Text1(0).Text & " " & KB.TextOutPut
End If
   KB.TextOutPutStyle = Basica
If Len(KB.TextOutPut) Then
Text1(1).Text = Text1(1).Text & " " & KB.TextOutPut
End If
End Sub


':) Ulli's VB Code Formatter V2.13.6 (17/09/2002 1:53:17 PM) 10 + 50 = 60 Lines
Private Sub Text1_Change(Index As Integer)

End Sub

