VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsKeyBoardPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Copyright 2002 Roger Gilchrist
'This class grew out of playing around with Nokia Ringtone Player by Ovidiu Daniel Diaconescu
'It seemed logical that a keyboard rather than text was the way to input at least some of the data
' to write tunes, so I wrote this.
'All you really need is a PictureBox.

'(Put it in a Frame and the Keyboard will make use of the Frame caption but this is not essential.)
'The keyboard is fitted to the frame shape so a short and wide one is best
'Place in a Declare section
'       Public KB As New ClsKeyBoardPicture
'Place in program's start up point
'       KB.AssignControl PictureBox1
'Place the class's matching Sub in the following Subs of the PictureBox1
'PictureBox1_MouseDown
''              KB.MouseDown                    draw pressed key, play a note
'PictureBox1_MouseMove
''              KB.MouseMove                    Send messages to Frame(if it exists)
'PictureBox1_MouseUp
''              KB.MouseUp                     'Unpress button
''
''Private MouseRead   private routine provides the data for the other Mouse events
''Private MouseDuration returns a duration base on button and shift
'    B1 = 16 B2 = 32 B4 = 64 Shift+B1 =2 Shift+B2= 4 Shift+B4=8  Ctrl = 1
'You now have a playable keyboard
''
'To link to other controls call the relevant Public Properties (see Class outputs below)

'Set various colours in the control     Default
''KB.BackColor                          VbBlack
''KB.ColorBlackKey                      VbBlack
''KB.ColorKeyEdge                       Grey       Key Edges are drawn to give non-responsive areas to allow proper note detection
''KB.UseCKeyColourTags                  True       Turn on the following two colour uses
''KB.WholeKeyTag                        False      Tag is bottom 1/10 of key ; True Whole Key is tag colours
''KB.ColorCKey                          VbYellow   Mark any C key
''KB.ColorMidCKey                       VBRed      Mark out Middle C Key
''KB.ColorPressed                       Grey       Key colour changes to this while a note is playing

''Kb.KeyBoardLook                                   This routine allows you to select stored colour schemes
''                                                  rather than setting each element b(feel free to add to the set
''                                                  Remember to update the KBLook Enum
''
''KB.BorderStyle                                   'Wrapper for PictureBox1.BorderStyle 0 Flat 1 raised
''



''KB.OffSetEnds                         True        Draw keyBoard with small Left and Right borders of BackColor

''KB.PauseKeyOn                         False       Hide Pause key. If you only want to play with keyboard a pause is not useful
'                                                   True a Pause Key allows mouse input of pauses if you are sending note text to textboxes
''KB.ColorPause                         Purple      The Pause Key is drawn one key to the right of the rest of the keyboard and distinctly different
'
''KB.UseFormCaption                     False       Allows keyboard to use the Form's caption to display messages

'Class outputs      These allow you to see and write the classes values to TextBoxes

''KB.DurationVal     Length of Note
''KB.Octave          Octave number
''KB.OctaveMember     [#]<A-G>
''KB.TextOutPutStyle                     Basica    Set the format of TextOutPut
''KB.TextOutPut                                    Emit a Printable Code String (Basica or Nokia) to textbox
''KB.Tempo                                120      Used by the internal PlayKBNote .Call this Property separately if you need to emit it to TextBox

' Other Public members DeBug tools
''DEBUG ONLY  (DO NOT USE EXCEPT FOR DEBUG PURPOSES)
''KB.MouseX    ' show X position in Picturebox1
''KeyNumber     count of White keys reveals m_KeyNumber used as part of internal calculation
''KB.Frequency ' Show Frequency that would be used (in case you get the keys out of sync with the array)
''             ' Also makes a nice addition to the Frame Message

'Private members
''DrawBoard                 Draw whole board whenever Colour or format or size changes
''Key                       Draw Key and its surrounding blackkeys
''GetRemoveRight            move Right of string to another str
''InvertSharp               reverse order of <A-G># to #<A-G>
''KeepInBoundsSimple        simple guard routine
''Playable                  a properly formed note exists
''PlayKBNote                play the note; calls to
''playbeep                  play sound
''playsleep                 play pause
''FrequencyFromNoteName     used by PlayKBNote to extract Frequency from Array

Option Explicit
Private M_KB As PictureBox               'Internal name of PictureBox
Attribute M_KB.VB_VarHelpID = -1
Private M_KBFrame As Frame              'Optional Internal Name of Frame
Private M_KBForm As Form                 'Holds PictureBox.Parent
Private m_ShowOnFormCaption As Boolean   'Allow use of Form.Caption Area Default False
Private m_ContainerIsFrame As Boolean   'Test if Frame exists befor  referencing it

Private KeyWidth As Integer             'Key Layout Descriptors
Private KeyHeight As Integer
Private BKeyWidth As Integer
Private BKeyHeight As Integer
Private KeyTabtop As Integer
Private KeyTabHeight As Integer
Private Const Def_KeyTabTopDiv As Integer = 8   'Divisors for ColourTag affects bottom 10% of keys
Private Const Def_KeyTabHeightDiv As Integer = 9

Private m_IsBlack As Boolean            'Detect which coloured Key the mouse is over
Private M_IsWhite As Boolean            'based on detecting the colour with M_KB.Point(X,Y)
Private M_IsPause As Boolean

Private m_WhitePressed As Boolean       'M_IsXXX AND Button>0 sets these
Private m_LeftBlackPressed As Boolean
Private m_RightBlackPressed As Boolean

Private m_Octave As Integer             'Store current octave
Private m_OctaveMember As Integer       'Store Current note name
Private PrevPressedOctave As Integer    'Store PreviousOctave while reading Pause key
Private m_KeyNumber As Long             'White key count; part of the determining which key you pressed(Not Directly meaningful)
Private m_Duration As Integer           'Store a duration value
Private m_UseColorTags As Boolean       'True =  colour tagged C keys
Private m_MouseX As Single              'Current X value of mouse cursor
Private m_Tempo As Integer              'Current Tempo Def 120

'Private colour values                  'These colours MUST BE different to each other
Private m_C_Colour As Long              'as they play a role in determining which key was clicked
Private m_MidC_Colour As Long           'Just use RGB() and very small differences if you want a
Private m_PressedColour As Long         'Mono-chromatic look
Private m_EdgeColour As Long
Private m_BackColour As Long
Private m_BlackKeyColour As Long
Private m_WhiteKeyColour As Long
Private m_PauseKeyColour As Long
Private m_WholeKeyTag As Boolean

Private m_OffSetEnds As Boolean     'True = small border of BackColor on Left and right of keyboard

Private KWLet As Variant '           'Stores A-G string values
Private KBLet As Variant             'Stores mathing sharps values

Private m_PauseKeyOn As Boolean      'True Show Pause Key False hide it(Keybard stretches to cover space created)
Private Const PauseKeyNo As Integer = 62 'Send this value to key if using PauseKey

Private Const KeyCount As Integer = 60   'Number of WhiteKeys



Private Type NoteValue            'Structure of NoteValueA array member
    Name As String
    Value As Single
End Type
Private NoteValueA() As NoteValue 'Array of Notes and Frequencies

Public Enum TextTypeOutPut ' Enum for TextOutPut
    Basica
    Nokia
End Enum
Private m_UseTxtOutPut As TextTypeOutPut 'Current TexpOutPut Style

Rem Mark Off
'Stops Code formatter complaining about these
#If False Then 'Enforce Case For Enums (does not compile but fools IDE)
Dim Basica
Dim Nokia
#End If  'Barry Garvin VBPJ 101 Tech Tips 11 March 2001 p1
Rem Mark On

Public Enum KBLook
    Antique
    Classical
    Default
End Enum

Rem Mark Off
#If False Then 'Enforce Case For Enums (does not compile but fools IDE)
Dim Antique
Dim Classical
Dim Default
#End If  'Barry Garvin VBPJ 101 Tech Tips 11 March 2001 p1

Dim ColourRGB
Dim ColourLong
Dim Description
Rem Mark On
'declare the beep and sleep functions private
'for use only inside the class
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Property Set AssignControl(ByVal PBox As PictureBox)

  'Copyright 2002 Roger Gilchrist
  'Set PBox and set up the keyboard to fit the PictureBox

    Set M_KB = PBox
    M_KB.AutoRedraw = True 'This is NOT to be changed; Keyboard will not display if you change this

    M_KB.BorderStyle = 1   'This can be changed through the wrapper Property

    'This is a non-essential pretty; If the PictureBox is in a frame then set it up to display messages
    'See MouseMove for mechanics of this
    Set M_KBForm = M_KB.Parent
    If TypeOf M_KB.Container Is Frame Then
        Set M_KBFrame = M_KB.Container
        m_ContainerIsFrame = True
        M_KBFrame.Caption = "Initialising keyboard"
        'M_KBFrame.BackColor = m_BackColour 'could be ugly
    End If

    DrawBoard

End Property

Public Property Get BackColor() As Long

    BackColor = m_BackColour

End Property

Public Property Let BackColor(ByVal Colr As Long)

    m_BackColour = Colr
    DrawBoard

End Property

Public Property Let BorderStyle(ByVal vNewValue As Variant)

  'wrapper Property for the PictureBox

    KeepInBoundsSimple 0, vNewValue, 1
    M_KB.BorderStyle = vNewValue

End Property

Public Property Get BorderStyle() As Variant

    BorderStyle = M_KB.BorderStyle

End Property

Private Sub Class_Initialize()

  'Copyright 2002 Roger Gilchrist
  'Note there are three different implementations of this Collection in this demo
  'each class has a different one tailored to its needs. You can not swap and change them.

  Dim TmpArrayValue As Variant, TmpArrayName As Variant, i As Integer

    'Default values
    m_Tempo = 120
    KWLet = Array("C", "D", "E", "F", "G", "A", "B", "P") 'Arrays used to name keys
    KBLet = Array("#C", "#D", "#F", "#F", "#G", "#A", "")
    m_UseColorTags = True          'Show C keys with colour tagging
    KeyBoardLook Default
    m_WholeKeyTag = False          'C key colour tags can be whole key or just lower tip
    m_OffSetEnds = True            ' ends of key board can be against left & right edges of PBox
    'or slightly inset (Looks better; helps MouseMove message to stay upto date)

    m_PauseKeyOn = False
    
    m_UseTxtOutPut = Basica
    m_ShowOnFormCaption = False

    'this collection is based on a colFR collection copyright of Barry Dune
    'The collection (and my array version of it) represent
    'the frequencies of piano keys and their notation in Nokia code
    'except for the first value "0", 32767 which is only there to support the Basica note code 'N0'
    'You could construct a similar collection for other musical notations
    'but remember Basica and (I guess) Nokia code are piano based.
    'I have laid the TmpArrayXXX out one octave to a line
    'to make them a bit easier to play with.
    '(see 'Note on NoteCollection.txt' for more details.

    'Initialize the returned key values

    TmpArrayValue = Array(32767, _
                    32.703, 34.648, 36.708, 38.891, 41.203, 43.654, 46.249, 48.999, 51.913, 55, 58.27, 61.735, _
                    65.406, 69.296, 73.416, 77.782, 82.407, 87.307, 92.499, 97.999, 103.826, 110, 116.541, 123.471, _
                    130.813, 138.591, 146.832, 155.564, 164.814, 174.614, 184.997, 195.998, 207.652, 220, 233.082, 246.942, _
                    261.626, 277.183, 293.665, 311.127, 329.628, 349.228, 369.994, 391.995, 415.305, 440, 466.164, 493.883, _
                    523.251, 554.365, 587.33, 622.254, 659.255, 698.457, 739.989, 783.991, 830.609, 880, 932.328, 987.767, _
                    1046.502, 1108.731, 1174.659, 1244.508, 1318.51, 1396.913, 1479.978, 1567.982, 1661.219, 1760, 1864.655, 1975.533, _
                    2093.005, 2217.461, 2349.318, 2489.016, 2637.021, 2793.826, 2959.956, 3135.964, 3322.438, 3520, 3729.31, 3951.066, _
                    4186.009, 4434.922, 4698.637, 4978.032, 5274.042, 5587.652, 5919.912, 6271.928, 6644.876, 7040, 7458.62, 7902.133, _
                    8372.019, 8869.845, 9397.273, 9956.064, 10548.083, 11175.305, 11839.823, 12543.855, 13289.752)

    TmpArrayName = Array("0", _
                   "C2", "#C2", "D2", "#D2", "E2", "F2", "#F2", "G2", "#G2", "A2", "#A2", "B2", _
                   "C3", "#C3", "D3", "#D3", "E3", "F3", "#F3", "G3", "#G3", "A3", "#A3", "B3", _
                   "C4", "#C4", "D4", "#D4", "E4", "F4", "#F4", "G4", "#G4", "A4", "#A4", "B4", _
                   "C5", "#C5", "D5", "#D5", "E5", "F5", "#F5", "G5", "#G5", "A5", "#A5", "B5", _
                   "C6", "#C6", "D6", "#D6", "E6", "F6", "#F6", "G6", "#G6", "A6", "#A6", "B6", _
                   "C7", "#C7", "D7", "#D7", "E7", "F7", "#F7", "G7", "#G7", "A7", "#A7", "B7", _
                   "C8", "#C8", "D8", "#D8", "E8", "F8", "#F8", "G8", "#G8", "A8", "#A8", "B8", _
                   "C9", "#C9", "D9", "#D9", "E9", "F9", "#F9", "G9", "#G9", "A9", "#A9", "B9", _
                   "C10", "#C10", "D10", "#D10", "E10", "F10", "#F10", "G10", "#G10")

    ReDim NoteValueA(UBound(TmpArrayValue) + 1) As NoteValue
    For i = LBound(TmpArrayValue) To UBound(TmpArrayValue)
        NoteValueA(i + 1).Name = TmpArrayName(i)
        NoteValueA(i + 1).Value = TmpArrayValue(i)
    Next i

End Sub

Public Property Let ColorBlackKey(ByVal Colr As Long)

    m_BlackKeyColour = Colr
    DrawBoard

End Property

Public Property Get ColorBlackKey() As Long

    ColorBlackKey = m_BlackKeyColour

End Property

Public Property Let ColorCKey(ByVal Colr As Long)

    m_C_Colour = Colr
    DrawBoard

End Property

Public Property Get ColorCKey() As Long

    ColorCKey = m_C_Colour

End Property

Public Property Let ColorKeyEdge(ByVal Colr As Long)

  'Copyright 2002 Roger Gilchrist
  'Set edge Colour; Any Appearance change should call DrawBoard

    m_EdgeColour = Colr
    DrawBoard

End Property

Public Property Get ColorKeyEdge() As Long

    ColorKeyEdge = m_EdgeColour

End Property

Public Property Let ColorMidCKey(ByVal Colr As Long)

  'Copyright 2002 Roger Gilchrist
  'Set MidC Colour; Any Appearance change should call DrawBoard

    m_MidC_Colour = Colr
    DrawBoard

End Property

Public Property Get ColorMidCKey() As Long

    ColorMidCKey = m_MidC_Colour

End Property

Public Property Let ColorPause(ByVal Colr As Long)

    m_PauseKeyColour = Colr
    DrawBoard

End Property

Public Property Get ColorPause() As Long

    ColorPause = m_PauseKeyColour

End Property

Public Property Let ColorPressed(ByVal Colr As Long)

  'Copyright 2002 Roger Gilchrist
  'Set edge Colour; Any Appearance change should call DrawBoard

    m_PressedColour = Colr
    DrawBoard

End Property

Public Property Get ColorPressed() As Long

    ColorPressed = m_PressedColour

End Property

Public Property Let ColorWhiteKey(ByVal Colr As Long)

  'Copyright 2002 Roger Gilchrist
  'Set White Key Colour; Any Appearance change should call DrawBoard

    m_WhiteKeyColour = Colr
    DrawBoard

End Property

Public Property Get ColorWhiteKey() As Long

    ColorPressed = m_WhiteKeyColour

End Property

Private Sub DrawBoard()

  'Copyright 2002 Roger Gilchrist
  'Draw whole board

  Dim i As Long


    'test that you have initializsed the class first
    If M_KB Is Nothing Then
        Exit Sub '>---> Bottom
    End If

    'Set basic measurements;
    'White Keys
    M_KB.BackColor = m_BackColour
    KeyWidth = M_KB.Width / IIf(m_PauseKeyOn, KeyCount + 2, KeyCount) + IIf(m_OffSetEnds, -3, 0)
    KeyHeight = M_KB.Height * 0.9
    'BlackKeys
    BKeyWidth = KeyWidth * 0.7
    BKeyHeight = KeyHeight * 0.7

    'C Key Tags; Change these through the WholeKeyTag or KeyTag property

    If m_WholeKeyTag Then
        KeyTabtop = KeyHeight
        KeyTabHeight = KeyHeight
      Else 'M_WHOLEKEYTAG = FALSE
        KeyTabtop = KeyHeight / Def_KeyTabTopDiv
        KeyTabHeight = KeyHeight / Def_KeyTabHeightDiv
    End If
    For i = 1 To KeyCount
        If i <> 61 Then
            Key i
        End If
    Next i
    If m_PauseKeyOn Then
        Key PauseKeyNo
    End If
    M_KB.Refresh

End Sub

Private Function Duration(NoteValue) As Long ', Dotted As Integer) As Long

  'plays a beep on frequency F durating V
  'Extracted from

  Dim Gama As Long

    On Error Resume Next
        Gama = 240000 / NoteValue
        Duration = Gama / m_Tempo
On Error GoTo 0
End Function

Public Property Get DurationVal() As Integer

  'Return Mouse button based Duration to program

    DurationVal = m_Duration

End Property

Public Function Frequency() As Single

  'Copyright 2002 Roger Gilchrist
  'Mostly for debug purposes

    Frequency = FrequencyFromNoteName(OctaveMember & (m_Octave + 2))

End Function

Private Function FrequencyFromNoteName(ByVal n$) As Single

  Dim i As Long

    For i = 1 To UBound(NoteValueA)
        If NoteValueA(i).Name = n$ Then
            FrequencyFromNoteName = NoteValueA(i).Value
            Exit Function '>---> Bottom
        End If
    Next i

End Function

Private Function GetRemoveRight(str As String, Optional CharNum As Long = 1) As String

  'Copyright 2002 Roger Gilchrist
  'take number of characters from right of str and remove them from str

    GetRemoveRight = Right$(str, CharNum)
    str = Left$(str, Len(str) - CharNum)

End Function

Private Function GetRemoveLeft(str As String, Optional CharNum As Long = 1) As String

  'Copyright 2002 Roger Gilchrist
  'take number of characters from right of str and remove them from str

    GetRemoveLeft = Left(str, CharNum)
    str = Mid$(str, 1 + CharNum)

End Function


Private Function InvertSharp(str As String) As String

    If InStr(str, "#") Then
        InvertSharp = Mid$(str, 2) & Left$(str, 1)
      Else 'NOT INSTR(STR,...
        InvertSharp = str
    End If

End Function

Private Sub KeepInBoundsSimple(Min As Variant, Val As Variant, MAx As Variant)

  'Copyright 2002 Roger Gilchrist
  'basic version of this sort of thing see ClsMusic_Nokia for complex version

    If Val < Min Then
        Val = Min
      ElseIf Val > MAx Then 'NOT VAL...
        Val = MAx
    End If

End Sub

Private Sub Key(KeyNo As Long)

  'Copyright 2002 Roger Gilchrist
  'Draw one key and surrounding black keys if necessary
  'Show Pressed colour if argument is true
  'NOTE because I developed this for a one note at a time player there is no allowance for chords
  'If you need it you could operate this by storing arrays of pressed states for each key LeftB, White and RightB
  'and use the arrays to trigger the IIf colour statements
  
  Dim OctaveMember As Integer 'Variables required in more than one place in Sub so set once call often
  Dim OctaveNo As Integer
  Dim KeyX1 As Integer
  Dim BOffSet As Integer
  Dim TopOffset As Integer  ' set top of keyboard just far enough from PictureBox top to allow

    ' the Off-Keyboard MouseMove message to appear
    If (m_PauseKeyOn = False And KeyNo > KeyCount) Or _
       (m_PauseKeyOn And KeyNo > 62 And KeyNo <> 63) Then 'Spaces around Pause key
        m_IsBlack = False
        M_IsWhite = False
        M_IsPause = False
        Exit Sub '>---> Bottom
    End If
    TopOffset = 9
    OctaveMember = KeyNo Mod 7 'Detect any C and decide whether to draw Black Keys
    OctaveNo = Int(KeyNo / 7) 'detect Middle C
    KeyX1 = (KeyWidth * KeyNo) - (KeyWidth) + IIf(m_OffSetEnds, 40, 0) 'initial start for drawing key
    BOffSet = (KeyWidth / 2)  'offset value for the Black keys

    M_KB.Line (KeyX1, TopOffset)-Step(KeyWidth, KeyHeight + TopOffset), IIf(m_WhitePressed, m_PressedColour, IIf(KeyNo = PauseKeyNo, m_PauseKeyColour, m_WhiteKeyColour)), BF

    If m_UseColorTags Then ' Colour tag system
        If Not m_WhitePressed Then 'don't tag if drawing a pressed key
            'You could expand this to give a separate color to each key(OctaveMember is a 0 to 6 range
            If OctaveMember = 0 Then 'If its a C then --------------------------------------------------- and if it is Middle C then
                M_KB.Line (KeyX1, KeyHeight - KeyTabtop + TopOffset)-Step(KeyWidth, KeyTabHeight), IIf(OctaveNo = 4, m_MidC_Colour, m_C_Colour), BF
            End If
        End If
    End If
    M_KB.Line (KeyX1, TopOffset)-Step(KeyWidth, KeyHeight + TopOffset), m_PressedColour, B 'Outline key
    'This draw routine hits the same black keys many times because they need to be redrawn every time they or the underlying white key is pressed
    If KeyNo < 62 Then
        KeyX1 = KeyX1 + BOffSet  'Right Black key
        Select Case OctaveMember
          Case 2, 6 'These don't have Sharps
          Case Else
            If KeyNo > 0 And KeyNo < KeyCount Then 'Not first or last key D0 and G10 which would normally have a Flat and a Sharp key
                'You could expand this to give a separate color to each key(OctaveMember is a 0 to 6 range
                M_KB.Line (KeyX1 + KeyWidth / 10, TopOffset)-Step(BKeyWidth, BKeyHeight), IIf(m_RightBlackPressed, m_PressedColour, m_BlackKeyColour), BF
                M_KB.Line (KeyX1 + KeyWidth / 10, TopOffset)-Step(BKeyWidth, BKeyHeight), m_PressedColour, B
            End If
        End Select
        KeyX1 = KeyX1 - (BOffSet * 2) 'Left Black key 'added one BoffSet so remove two
        Select Case OctaveMember
          Case 0, 3 'These don't have flats
          Case Else
            If KeyNo > 1 Then 'And ((KeyNo < KeyCount - 1) And m_PauseKeyOn) Then
                'You could expand this to give a separate color to each key(OctaveMember is a 0 to 6 range
                If m_PauseKeyOn And KeyNo = 62 Then

                  Else 'NOT M_PAUSEKEYON...
                    M_KB.Line (KeyX1 + KeyWidth / 10, TopOffset)-Step(BKeyWidth, BKeyHeight), IIf(m_LeftBlackPressed, m_PressedColour, m_BlackKeyColour), BF
                    M_KB.Line (KeyX1 + KeyWidth / 10, TopOffset)-Step(BKeyWidth, BKeyHeight), m_PressedColour, B
                End If
            End If
        End Select
        If m_ContainerIsFrame Then
            'M_KBFrame.BackColor = m_BackColour
            M_KBFrame.ForeColor = m_WhiteKeyColour 'does nothing in XPAware Programs
        End If
    End If
    M_KB.Refresh

End Sub

Public Sub KeyBoardLook(Look As KBLook)

    Select Case Look
      Case Antique
        ColorWhiteKey = RGB(255, 255, 230)
        ColorBlackKey = RGB(0, 0, 30)
        ColorCKey = RGB(255, 450, 50) '255, 245, 0)
        ColorMidCKey = RGB(255, 240, 100) '255, 210, 0)
        ColorPressed = RGB(155, 155, 130)
        ColorKeyEdge = m_PressedColour
        ColorPause = RGB(251, 251, 230)
        WholeKeyTag = True
      Case Classical
        ColorWhiteKey = RGB(255, 255, 250)
        ColorBlackKey = RGB(30, 30, 30)
        ColorCKey = RGB(255, 255, 245)
        ColorMidCKey = RGB(255, 255, 245)
        ColorPressed = RGB(210, 210, 185)
        ColorKeyEdge = m_PressedColour
        ColorPause = RGB(251, 251, 250)
      Case Default
       ColorPause = RGB(160, 160, 255)
        ColorCKey = vbYellow 'Default C key
        ColorMidCKey = vbRed  'Default Middle C
        ColorPressed = 15132390  'Default Key Down Colour
        ColorKeyEdge = m_PressedColour 'Key Edges ; Key edges are to allow a little lee way for the class
        'to work out which black key is being clicked
        m_BackColour = vbBlack         'PictureBox colour where not covered by keys
        ColorWhiteKey = vbWhite     '
        ColorBlackKey = vbBlack     '
        WholeKeyTag = False          'C key colour tags can be whole key or just lower tip
      Case Else
        KeyBoardLook Default 'recursive safety net
    End Select

End Sub

Public Sub MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'Copyright 2002 Roger Gilchrist

    MouseRead Button, Shift, X, Y 'Make sure it is up-to-date then
    If Playable Then
        Key m_KeyNumber
        PlayKBNote DurationVal & OctaveMember & Octave
    End If

End Sub

Sub ReadNokia(str$)
'Copyright 2002 Roger Gilchrist
'Play Nokia code
Dim Notes, i As Integer
'Nokia format <Duration>P
'Nokia Note Format <Duration>[#]<Note><Octave>
Notes = Split(Trim(str))
For i = LBound(Notes) To UBound(Notes)

PlayKBNote Notes(i)

Next

End Sub

Sub ReadBasica(str)
'Copyright 2002 Roger Gilchrist
'convert simplifed Basica to Nokia Code format
'Basica       "P<Duration>"
'Basica Note Format "L<Duration> O<Octave> [#]<Note>"
Dim Notes, i As Integer
Dim Dur As Integer, n As String, Oct As String
Dim Junk$
Dim Playme As String
Notes = Split(Trim(str))
For i = LBound(Notes) To UBound(Notes)
If Left(Notes(i), 1) = "P" Then
    Playme = Mid(Notes(i), 2) & "P"
Else
    Junk$ = Notes(i)
    Junk = Replace(Junk, "L", "")
    Junk = Replace(Junk, "O", " ")
    Dur = CInt(Left(Junk, InStr(Junk, " ")))
    Junk = Mid(Junk, InStr(Junk, " ") + 1)
    Oct = ""
    Do While IsNumeric(Left(Junk, 1))
    Oct = Oct & GetRemoveLeft(Junk)
    Loop
    Playme = Dur & InvertSharp(Junk) & CInt(Oct)
End If
PlayKBNote Playme

Next

End Sub
Public Sub MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'Copyright 2002 Roger Gilchrist
  'Read position then Display it on the frame(if it exists)or the Form if That is allowed
  'This wokrs best if there is a small border of non-keyboard PictureBox around the KeyBoard so that the Second message appears
  'when ever the mouse is off the keyboard (also appears if cursor is over the keyEdges).

  Dim Msg As String
  Static Prevmsg As String

    MouseRead Button, Shift, X, Y

    If Playable Then
        If OctaveMember = "P" Then
            Msg = "Keyboard (Pause Key  Duration: L = 16 R= 32 M = 64  Shift+L=2 Shift+R=4 Shift+M=8  Ctrl+Any=1)"
          Else 'NOT OCTAVEMEMBER...
            Msg = "Keyboard (Octave: " & Octave & "  Note:" & OctaveMember & "  Freq: " & Frequency & " )"
        End If
      Else 'PLAYABLE = FALSE
        Msg = "Keyboard (Mouse Button Note Duration: L = 16 R= 32 M = 64  Shift+L=2 Shift+R=4 Shift+M=8  Ctrl+Any=1)"
    End If

    If Msg <> Prevmsg Then
        If m_ContainerIsFrame Then
            M_KBFrame.Caption = Msg
          Else 'M_CONTAINERISFRAME = FALSE
            If m_ShowOnFormCaption Then
                M_KBForm.Caption = Msg
            End If
        End If
    End If
    Prevmsg = Msg

End Sub

Private Sub MouseRead(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'Copyright 2002 Roger Gilchrist
  'gather information on keyboard position and mouse button states; sets various variables ad Properties for rest of class

  Dim WhiteTest As Long
  Dim ColourTest As Long

    m_MouseX = X
    ColourTest = M_KB.Point(X, Y)
    m_IsBlack = (ColourTest = m_BlackKeyColour)
    M_IsWhite = (ColourTest = m_WhiteKeyColour) Or (ColourTest = m_C_Colour) Or (ColourTest = m_MidC_Colour)
    M_IsPause = (ColourTest = m_PauseKeyColour) And m_PauseKeyOn
    m_WhitePressed = M_IsWhite And Button > 0
    m_LeftBlackPressed = False
    m_RightBlackPressed = False
    m_KeyNumber = Int((X + KeyWidth) / KeyWidth)
    If m_PauseKeyOn Then

        If m_KeyNumber > KeyCount And m_KeyNumber <> 62 Then
            m_IsBlack = False
            M_IsWhite = False
            M_IsPause = False
            Exit Sub '>---> Bottom
        End If

    End If



        m_OctaveMember = (m_KeyNumber) Mod 7
        m_Duration = MouseDuration(Button, Shift)
    
    If Playable Then
        '  On Error Resume Next
        If m_IsBlack Then
            If X <= (KeyWidth * m_KeyNumber) - (KeyWidth / 2) Then
                m_OctaveMember = m_OctaveMember - 1
                m_LeftBlackPressed = Button > 0
              Else 'NOT X...
                m_OctaveMember = m_OctaveMember
                m_RightBlackPressed = Button > 0
            End If

        End If
        If M_IsPause Then
            m_Octave = PrevPressedOctave
          Else 'M_ISPAUSE = FALSE
            m_Octave = Int(m_KeyNumber / 7)
            If m_WhitePressed Or m_LeftBlackPressed Or m_RightBlackPressed Then
                PrevPressedOctave = m_Octave
            End If
        End If

    End If
    'On Error GoTo 0

    '    Constant (Button) Value Description
    'vbLeftButton 1 Left button is pressed.
    'vbRightButton 2 Right button is pressed.
    'vbMiddleButton 4 Middle button is pressed.
    '
    'Constant (Shift) Value Description
    'vbShiftMask 1 SHIFT key is pressed.
    'vbCtrlMask 2 CTRL key is pressed.
    'vbAltMask 4 ALT key is pressed.
    '
    '
    'The constants then act as bit masks you can use to test for any combination of buttons without having to figure out the unique bit field value for each combination.
    '
    'You test for a condition by first assigning each result to a temporary integer variable and then comparing the button or shift arguments to a bit mask. Use the And operator with each argument to test if the condition is greater than zero, indicating the key or button is pressed, as in this example:
    '
    'LeftDown = (Button And vbLeftButton) > 0
    'CtrlDown = (Shift And vbCtrlMask) > 0

End Sub

Private Function MouseDuration(Button As Integer, Shift As Integer)
    'vbLeftButton 1
    'vbRightButton 2
    'vbMiddleButton 4
    'vbShiftMask 1
    'vbCtrlMask 2
    'vbAltMask 4
'B1 = 16 B2 = 32 B4 = 64 Shift+B1 =2 Shift+B2= 4 Shift+B4=8  Ctrl = 1
MouseDuration = m_Duration
Select Case Button 'Base values
          Case vbLeftButton
            MouseDuration = 16
          Case vbRightButton
            MouseDuration = 32
          Case vbMiddleButton
            MouseDuration = 64
           Case Else
           Exit Function
End Select
If Shift Then
     If (Shift And vbShiftMask) Then ' Shift increase
            MouseDuration = MouseDuration / 8
     ElseIf (Shift And vbCtrlMask) Then
            MouseDuration = 1
     End If
End If
End Function

Public Sub MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'Copyright 2002 Roger Gilchrist
  'UnPress key

    If Playable Or M_KB.Point(X, Y) = m_PressedColour Then
        m_LeftBlackPressed = False
        m_RightBlackPressed = False
        m_WhitePressed = False
        Key m_KeyNumber
    End If

End Sub

Public Property Get MouseX() As Single

  'Copyright 2002 Roger Gilchrist
  'Mostly for Debug purposes

    MouseX = m_MouseX

End Property

Public Property Get Octave() As Integer

  'Copyright 2002 Roger Gilchrist
  'ReadOnly public interface

    Octave = m_Octave

End Property

Public Property Get OctaveMember() As String

  'Copyright 2002 Roger Gilchrist
  'ReadOnly public interface

    If Playable Then
        If m_OctaveMember > -1 Then
            If m_IsBlack Then
                OctaveMember = KBLet(m_OctaveMember)
              ElseIf M_IsWhite Then 'M_ISBLACK = FALSE
                OctaveMember = KWLet(m_OctaveMember)
              Else 'M_ISWHITE = FALSE
                OctaveMember = "P"
            End If
        End If
    End If

End Property

Public Property Let OffSetEnds(ByVal UseOffSet As Variant)

  'Copyright 2002 Roger Gilchrist
  'can be set to false but recommend leaving True it makes the MouseMove messageing work better

    m_OffSetEnds = UseOffSet

End Property

Public Property Get OffSetEnds() As Variant

    OffSetEnds = m_OffSetEnds

End Property

Public Property Let PauseKeyOn(ByVal OnOff As Boolean)

    m_PauseKeyOn = OnOff

End Property

Public Property Get PauseKeyOn() As Boolean

    PauseKeyOn = m_PauseKeyOn

End Property

Public Property Get Playable() As Boolean

  'Copyright 2002 Roger Gilchrist
  'ReadOnly The cursor is over a key so a note can be selected

    Playable = (m_IsBlack Or M_IsWhite Or M_IsPause)

End Property

Private Sub playbeep(F As Single, V As Long)

    KeepInBoundsSimple 36, F, 32768
    Beep F, V
    DoEvents

End Sub

Private Sub PlayKBNote(ByVal KeyName As String)
'Copyright 2002 Roger Gilchrist

  Dim Octave As Long
  Dim NoteName As String
  Dim NoteFreq As Single, DurationValue As Long
  Dim LOctave As Integer
  Dim DBUG As String

    DBUG = KeyName
        If InStr(KeyName, "P") Then
            DurationValue = CLng(Left$(KeyName$, InStr(KeyName$, "P") - 1))
            playsleep Duration(DurationValue)
        Else
    LOctave = CInt(GetRemoveRight(KeyName$)) + 2  'get the Octave value (usualy 1,2 or 3) and use it as an integer and '+ 4' convert it for the colFr
    NoteName = GetRemoveRight(KeyName$)  'get the note (c,d,e,f,g,a,b)
    If Right$(KeyName$, 1) = "#" Then 'if there is a "#"
        NoteName = GetRemoveRight(KeyName$) & NoteName
    End If
    NoteFreq = FrequencyFromNoteName(NoteName & LOctave)
    DurationValue = KeyName$
    playbeep NoteFreq, Duration(DurationValue) ', 0)    'play the note
    End If

End Sub

'plays a sleep (pause) durating V
Private Sub playsleep(d As Long)

    Sleep d
DoEvents
End Sub

Public Sub Resize()

  'Copyright 2002 Roger Gilchrist

    DrawBoard

End Sub

Public Property Get Tempo() As Integer

    Tempo = m_Tempo

End Property

Public Property Let Tempo(ByVal vNewValue As Integer)

    m_Tempo = vNewValue

End Property

Public Function TextOutPut() As String

  'Copyright 2002 Roger Gilchrist
If Len(OctaveMember) Then 'Prevents misfires printing out
    If m_UseTxtOutPut = Basica Then
        If OctaveMember = "P" Then
            'Basica       "P<Duration>"
            TextOutPut = OctaveMember & DurationVal
          Else 'NOT OCTAVEMEMBER...
            'Basica Note Format "L<Duration> O<Octave> [#]<Note>"
            TextOutPut = "L" & DurationVal & "O" & Octave & InvertSharp(OctaveMember)
        End If
      ElseIf m_UseTxtOutPut = Nokia Then 'NOT M_USETXTOUTPUT...

        If OctaveMember = "P" Then
            'Nokia format <Duration>P
            TextOutPut = DurationVal & OctaveMember
          Else 'NOT OCTAVEMEMBER...
            'Nokia Note Format <Duration>[#]<Note><Octave>
            TextOutPut = DurationVal & OctaveMember & Octave
        End If

    End If
End If
End Function

Public Property Get TextOutPutStyle() As TextTypeOutPut

    TextOutPutStyle = m_UseTxtOutPut

End Property

Public Property Let TextOutPutStyle(ByVal vNewValue As TextTypeOutPut)

    m_UseTxtOutPut = vNewValue

End Property

Public Property Get UseCKeyColourTags() As Boolean

    UseCKeyColourTags = m_UseColorTags

End Property

Public Property Let UseCKeyColourTags(ByVal UseCTag As Boolean)

  'Copyright 2002 Roger Gilchrist
  'Allow/Prevent Colour tagging; Any Appearance change should call DrawBoard

    m_UseColorTags = UseCTag
    DrawBoard

End Property

Public Property Get UseFormCaption() As Boolean

    UseFormCaption = m_ShowOnFormCaption

End Property

Public Property Let UseFormCaption(ByVal vNewValue As Boolean)

    m_ShowOnFormCaption = vNewValue

End Property

Public Property Get WholeKeyTag() As Boolean

    WholeKeyTag = m_WholeKeyTag

End Property

Public Property Let WholeKeyTag(ByVal WholeKey As Boolean)

  'Copyright 2002 Roger Gilchrist
  'Set TagKeyStyle; Any Appearance change should call DrawBoard

    m_WholeKeyTag = WholeKey
    DrawBoard

End Property

':) Ulli's VB Code Formatter V2.13.6 (17/09/2002 9:11:30 PM) 179 + 862 = 1041 Lines
