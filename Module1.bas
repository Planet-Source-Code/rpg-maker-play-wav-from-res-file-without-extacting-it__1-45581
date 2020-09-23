Attribute VB_Name = "res"
Option Explicit
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Private Const SND_SYNC = &H0        ' Play synchronously (default).
Private Const SND_NODEFAULT = &H2    ' Do not use default sound.
Private Const SND_MEMORY = &H4
Private Const SND_LOOP = &H8         ' Loop the sound until next
Private Const SND_NOSTOP = &H10      ' Do not stop any currently
Private Const SND_ASYNC = &H1          '  play asynchronously
Private bytSound() As Byte ' Always store binary data in byte arrays!
Public Enum SoundFlags
    soundSYNC = SND_SYNC
    soundNO_DEFAULT = SND_NODEFAULT
    soundMEMORY = SND_MEMORY
    soundLOOP = SND_LOOP
    soundNO_STOP = SND_NOSTOP
    soundASYNC = SND_ASYNC
End Enum

Public Enum AppSounds
    appsoundNT_LOGON_WAVE = 100
End Enum
Public Sub PlayWaveRes(vntResourceID As AppSounds, Optional vntFlags As SoundFlags = soundASYNC)
    bytSound = LoadResData(vntResourceID, "WAVE")
    If IsMissing(vntFlags) Then
        vntFlags = SND_NODEFAULT Or SND_SYNC Or SND_MEMORY
    End If
    If (vntFlags And SND_MEMORY) = 0 Then
        vntFlags = vntFlags Or SND_MEMORY
    End If
    sndPlaySound bytSound(0), vntFlags
End Sub
