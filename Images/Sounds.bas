Attribute VB_Name = "Sounds"
Option Explicit
Public Declare PtrSafe Function BeepAPI Lib "kernel32" Alias "Beep" _
             (ByVal Frequency As Long, ByVal Milliseconds As Long) As Long

Sub TestBeeper()
    'run this to test beeper
    
    Dim nFreq As Long, nDur As Long

    nFreq = 800     'frequency (Hertz)
    nDur = 500   'duration (milliseconds)

    'call beeper function
    Beeper nFreq, nDur

End Sub

Function Beeper(nF As Long, nD As Long) As Boolean
    'makes a beep sound of selected frequency and duration
    'This works for NT/2000/XP and beyond.
    'Before that, frequency and duration ignored.
    
    BeepAPI nF, nD

    Beeper = True

End Function

Sub TestNotes()
    'music notes played using known frequencies and those
    'calculated from knowlege of their relative positions'
        
    Dim vAN As Variant, vOTJ As Variant, vOTJD As Variant
    Dim i As Long, nFreq As Long
    Dim nDur As Long, nLen As Long

    'sets the basic note duration
    nDur = 500

    'store the specific frequencies in array - zero based...
    'these frequencies are for do,re,me,fa,so,la,te,do based on middle C =261.63Hz (262)
    'CDEFGABC
    vAN = Array(262, 294, 330, 349, 392, 440, 494, 523)
    
    'or store a jingle in note difference notation. Ode To Joy on beeper.
    vOTJ = Array(7, 7, 8, 10, 10, 8, 7, 5, 3, 3, 5, 7, 7, 5, 5) 'note positions from 440Hz
    vOTJD = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2)  'durations
    
    'scales up
    'CDEFGABC
    'do re me fa so la te do
    For i = 0 To 7
        nFreq = vAN(i)
        BeepAPI nFreq, nDur
    Next i

    Delay 1000 'delay one second

    'scales down
    'CBAGFEDC
    'do te la so fa me re do
    For i = 7 To 0 Step -1
        nFreq = vAN(i)
        BeepAPI nFreq, nDur
    Next i

    Delay 1000 'delay one second

    '34 notes, naturals, sharps and flats
    'played using note position from 440Hz
    For i = -5 To 28
        nFreq = CInt(440 * 2 ^ (i / 12))
        BeepAPI nFreq, nDur
    Next i

    Delay 1000 'delay one second

    'Ode to Joy - albeit crude, using note distance only
    For i = 0 To 14
       nFreq = CInt(440 * 2 ^ (vOTJ(i) / 12))
       BeepAPI nFreq, 400 * vOTJD(i)
    Next i
   
   Delay 1000 'delay one second
   
   'or use direct entries to make a custom sound
    BeepAPI 262 * 2, 200
    BeepAPI 494 * 2, 200
    BeepAPI 494 * 2, 200
    BeepAPI 262 * 2, 500
    
End Sub



Sub testSendMorse()
    'run this to test the Morse code sender
    'integers and simple alphabet only
    
    Dim sIn As String
    Dim start As Single, ends As Single
    
    sIn = "The quick brown fox jumps over the lazy dog 0123456789 times"
    
    'start = Timer
    SendMorse sIn, 440, 120 'string,freq (Hz),dot length (mS)

    'ends = Timer - start
    'MsgBox ends
End Sub

Sub SendMorse(ByVal sIn As String, nUF As Single, nUL As Single)
    'Sounds out Morse code for input string sIn
    'Parmeters frequency(Hz) and dot length(mS)

    Dim vSL As Variant, vSN As Variant, vM As Variant
    Dim i As Long, j As Long, nAsc As Integer
    Dim sWord As String, sCode As String

    'check that there is a decent string input
    If Trim(sIn) = "" Then
        MsgBox "Illegal characters in input string - closing"
        Exit Sub
    End If
        
    'load letter array with morse code- 1 for dot and 3 for dah
    vSL = Array("13", "3111", "3131", "311", "1", "1131", "331", "1111", "11", _
                "1333", "313", "1311", "33", "31", "333", "1331", "3313", "131", _
                "111", "3", "113", "1113", "133", "3113", "3133", "3311") 'a,b,c,...z
    'load number array with morse code- 1 for dot and 3 for dah
    vSN = Array("33333", "13333", "11333", "11133", "11113", _
                "11111", "31111", "33111", "33311", "33331")              '0,1,2,...9
        
    'split the input string into words
    vM = Split(Trim(sIn), " ") 'zero based
    
    For i = LBound(vM) To UBound(vM) 'step through words
        'get one word at a time
        sWord = LCase(vM(i)) 'current word
        'get one chara at a time
        For j = 1 To Len(sWord)
            'look up chara asci code
            nAsc = Asc(Mid(sWord, j, 1))
            'get morse sequence from array
            Select Case nAsc
                Case 97 To 122 'a letter
                    sCode = vSL(nAsc - 97)
                    MakeBeeps sCode, nUL, nUF
                    If j <> Len(sWord) Then
                        Delay (nUL * 3) 'add 3 spaces between letters
                    End If
                Case 48 To 57  'an integer
                    sCode = vSN(nAsc - 48)
                    MakeBeeps sCode, nUL, nUF
                    If j <> Len(sWord) Then
                        Delay (nUL * 3) 'add 3 spaces between letters
                    End If
                Case Else
                    MsgBox "Illegal character in input" & vbCrLf & _
                           "Only A-Z and 0-9 permitted."
            End Select
            
        Next j
        If i <> UBound(vM) Then Delay (nUL * 7) 'add 7 spaces between words
    Next i
    
End Sub
Function MakeBeeps(ByVal sIn As String, ByVal nUL As Single, ByVal nUF As Single) As Boolean
    'makes beep sounds for one character based on coded input string
    
    Dim i As Long, j As Long, nLen As Long
    Dim nT As Single, nE As Single
    
    For i = 1 To Len(sIn)
        'get character element
        nLen = CInt(Mid(sIn, i, 1))
        Select Case nLen
        Case 1
            BeepAPI nUF, nUL + Random(nUL)
            If i <> Len(sIn) Then Delay nUL
        Case 3
            BeepAPI nUF, (3 * nUL) + Random(3 * nUL)
            If i <> Len(sIn) Then Delay nUL
        Case Else
            MsgBox "error"
        End Select
    Next i
            
    MakeBeeps = True

End Function

Function Random(nDot As Single) As Single
    'adds a random variation to the timing
    'used to better hide machine code signature
    'nRand = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
    
    Dim nRand As Long, nPercent As Single
    Dim nU As Long, nL As Long
    
    'set a number here for max error percentage
    'eg; 10 for ten percent error, 0 for none.
    nPercent = 10 'max percent plus and minus

    'initialize the random generator
    Randomize
    
    'generate small random number as the timing error
    
    nRand = Int((nDot * nPercent / 100 + nDot * nPercent / 100 + 1) * Rnd - nDot * nPercent / 100)

    Random = nRand

End Function

Sub Delay(nD As Single)
    'delays for nD milliseconds
    'randomness set in Random()
    
    Dim start As Single
   
    nD = nD + Random(nD) 'add randomness to intention
    
    start = Timer  ' Set start time.
    Do While Timer < start + nD / 1000
        DoEvents    ' Yield to other processes.
    Loop

End Sub








  
   'or use direct entries to make a custom sound
    BeepAPI 262 * 2, 500
    Delay (500)
    BeepAPI 494 * 2, 1000
        Delay (500)
    BeepAPI 494, 500
        Delay (500)
    BeepAPI 262, 300
    
    
End Sub
