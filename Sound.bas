Attribute VB_Name = "Sound"
Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_SYNC = &H0         '  play synchronously (default)
Public Const SND_ASYNC = &H1         '  play asynchronously
Public Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound

Public Sub PlayTheWav(TheWav As String)
Dim SYNC As Long
Dim wav As String

wav = App.Path & "\" & TheWav

If wav = "" Then Exit Sub

    SYNC = SND_ASYNC

Dim R As Integer
R = sndPlaySound(ByVal wav, SYNC)
End Sub

Public Sub PlayTheWavLoop(TheWav As String)
Dim SYNC As Long
Dim wav As String

wav = App.Path & "\" & TheWav

If wav = "" Then Exit Sub

    SYNC = SND_LOOP

Dim R As Integer
R = sndPlaySound(ByVal wav, SYNC)
End Sub
