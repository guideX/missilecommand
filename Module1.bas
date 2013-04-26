Attribute VB_Name = "Module1"
Option Explicit

Declare Function GetPrivateProfileStringByKeyName& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey$, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Declare Function GetPrivateProfileStringKeys& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Declare Function GetPrivateProfileStringSections& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName&, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Declare Function WritePrivateProfileStringByKeyName& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String)
Declare Function WritePrivateProfileStringToDeleteKey& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Long, ByVal lplFileName As String)
Declare Function WritePrivateProfileStringToDeleteSection& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Long, ByVal lpString As Long, ByVal lplFileName As String)
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Public Type pType
        xs     As Single  'x coor source
        ys     As Single  'y coor source
        x      As Single  'x coor current
        Y      As Single  'y coor current
        xm     As Single  'x movemoment add
        ym     As Single  'y movement add
        xe     As Integer 'explode when reach this x (used to re-align missle)
        ye     As Integer 'explode when reach this y
        Link   As Integer 'pointer to next item
        Status As Integer '1=in flight, >=2 = size of explosion
        Smart  As Boolean
        SplitY As Integer '0=No, other = ycoord at which to split
End Type

Public Type LevelType
        LevelNo As Integer
        bMax   As Integer
        bDrop  As Integer
        bSpeed As Single
        mMax   As Integer
        mFire  As Integer
        mSpeed As Single
        Smart  As Integer   '% of smart bombs
        Split  As Integer   '% of split bombs
        Name As String
End Type
'||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

Public Const tFile = "build1.gif"

Public Const ShowM = 1, HideM = 2, ReadM = 3

Public Const bMax = 500
Public Const mMax = 500


Public Const MaxLevel = 10

Public Const MaxX% = 1000
Public Const MaxY% = 750

'||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

Public blnAllowResize  As Boolean
Public gstrPath As String

Public bClr As Long    'bomb
Public mClr As Long    'missle
Public lClr As Long    'line
Public eClr As Long    'erase

Public blnQuit As Boolean
Public blnFirstTime  As Boolean

Public Level(1 To MaxLevel) As LevelType

Public Targets() As Boolean

Public b(1 To bMax) As pType
Public m(1 To mMax) As pType

'//////////////////////////////////////////////
'ini stuff
Public SyncDelay As Single
Public SyncDist As Single, SyncTime As Single
Public l(1 To 10) As String 'level strings
Public SoundFX As Boolean
Public MaxTarget As Integer
Public bMaxStatus As Integer
Public mMaxStatus As Integer
Public mSpeed As Single
Public bExplodeb As Boolean
'//////////////////////////////////////////////

Public mMouseX As Single
Public mMouseY As Single
Public mMouseButton As Integer


Public Sub Alarm()

    If SoundFX Then sndPlaySound gstrPath & "Alarm.wav", 1

End Sub

Public Function x2t(x As Integer) As Integer

Dim i As Integer

    i = Int((UBound(Targets) * x) / MaxX%) + 1
    If i <= 0 Then i = 1
    If i > UBound(Targets) Then i = UBound(Targets)

    x2t = i

End Function



Public Sub MachineSpeed(SyncFactor)


'Dim SyncDist As Single
'Dim SyncTime As Single
'
''SyncDelay is machine specific
''smaller = better movement resolutuin but with lower max load
''to set timing on a particular machine:
''determine max load (ex: 50 in flight)
''  NOTE: Max load is level dependent!
''find SyncDelay that gives same time for min load as max load
''  i.e., set SyncDelay to one loop time at max load
''find speed that give desired dist/sec at that SyncDelay
''  this is defined by the formula Factor = ((Dist * Delay) / Time)
'
'SyncDelay = 0.1
'
''these settings give a nice speed, 2.5 seconds to fire from
''(320,640) to (0,0) i.e., origin to upper left hand corner
'
'SyncDist = 576
'SyncTime = 2

    SyncFactor = (SyncDist * SyncDelay) / SyncTime

End Sub

Public Sub InitLevels(Level() As LevelType, SyncFactor)


Dim avl As Variant
Dim i%
    
    
    
    For i% = 1 To MaxLevel

        avl = Split(l(i%), ",")

        Level(i%).LevelNo = i%
        
        Level(i%).bMax = avl(0) 'total # of bombs for this level
        Level(i%).mMax = avl(1) 'total # of missles for this level

        Level(i%).bDrop = avl(2) 'max inflight bombs (this level)
        Level(i%).mFire = avl(3) 'max inflight missles (this level)

        Level(i%).bSpeed = avl(4) 'relative speed of bombs

        Level(i%).Smart = avl(5) '% of total bombs that are smart (evade missles)
        Level(i%).Split = avl(6) '% of total bombs that have multiple warheads (splitters)

        Level(i%).Name = avl(7) 'the name of the level

        Level(i%).bSpeed = Level(i%).bSpeed * SyncFactor 'speed sync stuff
        Level(i%).mSpeed = mSpeed * SyncFactor

    Next i%


'these are the defaults, they work pretty good...

'        'bMax, mMax, bDrop, mFire, bSpeed, mSpeed (in 1/10th pixel per loop), Smart%, Split%, Name
'        Select Case i%
'            Case 1
'                L = Array(10, 50, 5, 5, 0.05, 1, 0, 15, "Slow and Dumb I")
'            Case 2
'                L = Array(10, 50, 5, 10, 0.05, 1, 0, 20, "Slow and Dumb II")
'            Case 3
'                L = Array(10, 50, 5, 10, 0.07, 1, 30, 25, "Faster and Smarter I")
'            Case 4
'                L = Array(15, 100, 10, 20, 0.09, 1, 40, 25, "Faster and Smarter II")
'            Case 5
'                L = Array(15, 100, 10, 20, 0.09, 1, 50, 25, "Faster and Smarter III")
'            Case 6
'                L = Array(20, 200, 10, 20, 0.12, 1, 50, 33, "Prelude")
'            Case 7
'                L = Array(30, 200, 10, 20, 0.12, 1, 60, 33, "Dooms Day I")
'            Case 8
'                L = Array(40, 200, 20, 30, 0.12, 1, 70, 33, "Dooms Day II")
'            Case 9
'                L = Array(50, 200, 50, 50, 0.12, 1, 80, 33, "You've got to be kidding!")
'            Case Else
'                L = Array(100, 200, 50, 50, 0.15, 1, 90, 50, "This ain't right")
'        End Select


End Sub

Public Sub BuildPool(Pool%, Head%, p() As pType, Max%)

Dim i%

    Head% = 0
    Pool% = 1
    For i% = 1 To Max%
        Call Nullp(p(i%))

        If i% <> Max% Then
            p(i%).Link = i% + 1
        Else
            p(i%).Link = 0
        End If
    Next i%

End Sub

Public Sub ResetTargets(t() As Boolean, f As String, p As PictureBox)
    
Dim Wide As Long
Dim High As Long
Dim x As Long, Y As Long
Dim AspectX As Long
Dim AspectY As Long
Dim i As Integer
Dim n As Integer
    
    
    n = UBound(t)
        
    FrmDD.Pic.BackColor = eClr
        
    FrmDD.Image1.Visible = True
    FrmDD.Image1.Stretch = False
    FrmDD.Image1.Picture = LoadPicture(f)
    
    AspectX = p.Width / n
    AspectY = (FrmDD.Image1.Height * AspectX) / FrmDD.Image1.Width
    
    FrmDD.Image1.Picture = Nothing
    
    FrmDD.Image1.Move 0, 0, AspectX, AspectY
    FrmDD.Image1.Stretch = True
    FrmDD.Image1.Picture = LoadPicture(f)
    
    p.Left = FrmDD.Image1.Width + 720
    
    DoEvents

    'Y = (p.Height - AspectY) * (1 / Screen.TwipsPerPixelX) + 10
    Y = (p.Height - AspectY) * (1 / Screen.TwipsPerPixelX)

    p.ScaleMode = vbPixels
    p.ScaleMode = vbPixels

    Wide = FrmDD.Image1.Width * (1 / Screen.TwipsPerPixelX)
    High = FrmDD.Image1.Height * (1 / Screen.TwipsPerPixelY)

    x = 0
    For i = 1 To n
        If t(i) Then
            BitBlt p.hDC, x, Y, Wide, High, p.Parent.hDC, 0, 0, &HCC0020
        End If
        x = x + Wide
    Next i
    
    p.Left = 30

    p.ScaleMode = vbTwips
    FrmDD.Image1.Picture = Nothing
    FrmDD.Image1.Visible = False
    
    FrmDD.Pic.ScaleWidth = MaxX%
    FrmDD.Pic.ScaleHeight = MaxY%
    
    
    DoEvents

End Sub

Public Sub MouseRead(x%, Y%, blnLeft, blnRight)

    blnLeft = False
    blnRight = False
    
    If mMouseButton <> 0 Then

        x% = mMouseX
        Y% = mMouseY
        
        If mMouseButton = 1 Then
            blnLeft = True
        Else
            'blnRight = True
        End If

        mMouseButton = 0

    End If

End Sub

Public Sub Delay(d#)

'TODO:

Dim t As Single

    t = Timer
    Do Until Timer - t > d#: Loop

End Sub

Public Sub FireFX()
    
    If SoundFX Then sndPlaySound gstrPath & "Swoosh.wav", 1

End Sub


Public Function InRange%(x1, y1, r%, x2, y2)

    InRange% = False
    If (x2 >= (x1 - r%)) And (x2 <= (x1 + r%)) Then
        If (y2 >= (y1 - r%)) And (y2 <= (y1 + r%)) Then
            InRange% = True
        End If
    End If

End Function

Public Sub Intercept(mHead%, m() As pType, bHead%, b() As pType, Level As LevelType)

Dim mLink%, bLink%

    mLink% = mHead%

    Do While mLink% <> 0
        
        If m(mLink%).Status <> 1 Then 'exploding

            'see if any in b's are within area of explosion

            bLink% = bHead%
            Do While bLink% <> 0
                If b(bLink%).Status = 1 Then 'in flight
                    If InRange%(m(mLink%).x, m(mLink%).Y, m(mLink%).Status, b(bLink%).x, b(bLink%).Y) Then
                        If SoundFX Then sndPlaySound gstrPath & "Thunder.wav", 1
                        b(bLink%).Status = 2
                    ElseIf b(bLink%).Smart Then   'take evasive action
                        
                        If InRange%(m(mLink%).x, m(mLink%).Y, mMaxStatus * 2, b(bLink%).x, b(bLink%).Y) Then


                            'evasive action
                            If b(bLink%).Y < m(mLink%).Y Then 'ex below us
                                If b(bLink%).xm > 0 Then
                                    If b(bLink%).x < m(mLink%).x Then
                                        b(bLink%).xm = -b(bLink%).xm
                                    End If
                                Else
                                    If (b(bLink%).x > m(mLink%).x) Then
                                        b(bLink%).xm = -b(bLink%).xm
                                    End If
                                End If
                            End If

                        End If
                    End If
                End If
                bLink% = b(bLink%).Link
            Loop

        End If
        mLink% = m(mLink%).Link
    Loop


End Sub

Public Sub LaunchB(Pool%, Head%, p() As pType, Level As LevelType, bDroped%, Targets() As Boolean, Optional Split As Boolean = False, Optional sx As Single = 0, Optional sy As Single = 0)

Dim tx%, x%, t%
Dim a As Single, s As Single

    If Pool% <> 0 And ((bDroped% < Level.bMax) Or Split) Then

        Call PickTarget(Targets(), tx%, x%)
        
        If Split Then 'override random pick
            x% = sx
        End If

        bDroped% = bDroped% + 1
        FrmDD.Label1(0) = Level.bMax - bDroped%

        
        t% = Pool%
        Pool% = p(t%).Link
        p(t%).Link = Head%
        Head% = t%

        
        
        If Split Then
            p(t%).x = sx    'split x
            p(t%).Y = sy    'split y
        Else
            p(t%).x = x%
            p(t%).Y = 0
        End If
        
        p(t%).xs = p(t%).x
        p(t%).ys = p(t%).Y
        
        p(t%).xe = tx%
        p(t%).ye = MaxY%
        
        If Level.Smart <> 0 Then
            p(t%).Smart = PickRandom%(1, 100) <= Level.Smart
        Else
            p(t%).Smart = 0
        End If
        
        p(t%).SplitY = 0
        If Level.Split <> 0 Then
            If PickRandom%(1, 100) <= Level.Split Then
                If p(t%).Y < MaxY% * 0.8 Then
                    p(t%).SplitY = PickRandom%(CInt(p(t%).Y), MaxY% * 0.8)
                
                    'Debug.Assert False
                
                End If
                'If p(t%).Y >= p(t%).SplitY Then 'past our split point
                '    p(t%).SplitY = 0
                'End If
            End If
        End If
        
        'Debug.Assert p(t%).SplitY <> 0
        
        p(t%).Status = 1

        If x% = tx% Then 'straight drop down
            p(t%).xm = 0
            p(t%).ym = Level.bSpeed + (Level.bSpeed * Rnd)
        Else
            'a = Atn(MaxY% / (x% - tx%))
            a = Atn((MaxY% - p(t%).Y) / (x% - tx%)) 'to accomadate splits in mid air
            p(t%).xm = Cos(a)
            p(t%).ym = Abs(Sin(a))
            If a > 0 Then p(t%).xm = -p(t%).xm

            s = Level.bSpeed + (Level.bSpeed * Rnd)
                                                    
            p(t%).xm = p(t%).xm * s
            p(t%).ym = p(t%).ym * s

        End If

    End If

End Sub

Public Sub LaunchM(x%, Y%, Pool%, Head%, p() As pType, Level As LevelType, mFired%)

Dim t%, xs%
Dim a As Single

    If Pool% <> 0 And mFired% < Level.mMax Then

        Call FireFX

        mFired% = mFired% + 1
        FrmDD.Label1(1) = Level.mMax - mFired%

        t% = Pool%
        Pool% = p(t%).Link
        p(t%).Link = Head%
        Head% = t%

        xs% = MaxX% / 2

        p(t%).x = xs%
        p(t%).Y = MaxY%
        
        p(t%).xs = p(t%).x
        p(t%).ys = p(t%).Y
        
        p(t%).xe = x%
        p(t%).ye = Y%
        p(t%).Status = 1

        If x% = xs% Then
            p(t%).xm = 0
            p(t%).ym = -Level.mSpeed
        Else
            a = Atn((MaxY% - Y%) / (x% - xs%))
            p(t%).xm = Cos(a)
            p(t%).ym = Abs(Sin(a))
            If a < 0 Then p(t%).xm = -p(t%).xm

            p(t%).xm = p(t%).xm * Level.mSpeed
            p(t%).ym = -(p(t%).ym * Level.mSpeed)

        End If

    Else
        If SoundFX Then sndPlaySound gstrPath & "Empty.wav", 1
    End If

End Sub

Public Sub Nullp(p As pType)

        p.xs = 0
        p.ys = 0
        p.x = 0
        p.Y = 0
        p.xm = 0
        p.ym = 0
        p.ye = 0
        p.Status = 0
        p.Smart = False
        p.SplitY = 0


End Sub

Public Function PickRandom%(l%, u%)

Dim t As Single

    If l% < u% And l% <= 1000 And u% >= 0 Then
        Do
            t = Rnd * 1000
        Loop Until t >= l% And t <= u%
    Else
        t = l%
    End If
    
    PickRandom% = t

End Function

Public Sub PickTarget(Targets() As Boolean, tx%, x%)

    x% = PickRandom(0, MaxX%)
    
    'make sure tx is on a remaining target

    Do
        tx% = PickRandom(0, MaxX%)
        DoEvents
    Loop Until Targets(x2t(tx%))

End Sub


Public Sub MyShow(Head%, Pool%, p() As pType, ByVal Clr&, ByVal MaxStat%, Droped%, Level As LevelType)

Dim pLink%, Link%, t%, tmpMaxStat%

'Dim NormalClr&, SplitClr&

    pLink% = 0
    Link% = Head%

    'NormalClr& = Clr&
    'SplitClr& = vbGreen

    Do While Link% <> 0

        If p(Link%).Status = 1 Then 'in flight
        
            'Clr& = NormalClr&
            'If p(Link%).SplitY <> 0 Then
            '    Clr& = SplitClr
            'End If

            FrmDD.Pic.Circle (p(Link%).x, p(Link%).Y), 2, eClr
            If Not p(Link%).Smart Then
                FrmDD.Pic.Line (p(Link%).xs, p(Link%).ys)-(p(Link%).x, p(Link%).Y), eClr
            End If

            p(Link%).x = p(Link%).x + p(Link%).xm
            p(Link%).Y = p(Link%).Y + p(Link%).ym

'force explosion to be where user intended it to be dispite sync error
If p(Link%).ym < 0 Then 'missle
    If p(Link%).Y <= p(Link%).ye Then
        p(Link%).x = p(Link%).xe
        p(Link%).Y = p(Link%).ye
    End If
End If

            FrmDD.Pic.Circle (p(Link%).x, p(Link%).Y), 2, Clr&
            
            If Not p(Link%).Smart Then
                FrmDD.Pic.Line (p(Link%).xs, p(Link%).ys)-(p(Link%).x, p(Link%).Y), lClr
            End If

            'test for explode here
            If p(Link%).ym < 0 Then 'missle
                If p(Link%).Y <= p(Link%).ye Then
                    
                    If SoundFX Then sndPlaySound gstrPath & "Explode.wav", 1
                    
                    p(Link%).Status = 2
                End If
            
            Else 'bomb
                
                If p(Link%).Y >= p(Link%).ye Then
                    
                    If SoundFX Then sndPlaySound gstrPath & "Thunder.wav", 1
                    
                    p(Link%).Status = 2
                    
                    'turn off target
                    Targets(x2t(CInt(p(Link%).x))) = False
                    
                Else 'test for off screen
                   
                   If p(Link%).x < 0 Or p(Link%).x > MaxX% Then
                        
                        FrmDD.Pic.Circle (p(Link%).x, p(Link%).Y), 2, eClr
                        
                        If Not p(Link%).Smart Then
                            FrmDD.Pic.Line (p(Link%).xs, p(Link%).ys)-(p(Link%).x, p(Link%).Y), eClr
                        End If
                        
                        GoSub RemoveIt
                   
                   Else 'test for split
                    
                        If p(Link%).SplitY <> 0 Then
                            If p(Link%).Y >= p(Link%).SplitY Then 'split
                            
                                If SoundFX Then sndPlaySound gstrPath & "Split.wav", 1
                                
                                p(Link%).SplitY = 0
                                
                                If Pool% = 0 Then 'add one to the pool!
                                    If Level.bMax < UBound(p) Then
                                        Level.bMax = Level.bMax + 1
                                        Pool% = Level.bMax
                                        Nullp p(Pool%)
                                        p(Pool%).Link = 0
                                    End If
                                End If
                                
                                LaunchB Pool%, Head%, p(), Level, Droped%, Targets(), True, p(Link%).x, p(Link%).Y
                            
                            End If
                        End If
                   
                   End If
                
                End If
            End If

        Else 'exploding

            FrmDD.Pic.Circle (p(Link%).x, p(Link%).Y), p(Link%).Status, eClr
            If Not p(Link%).Smart Then
                FrmDD.Pic.Line (p(Link%).xs, p(Link%).ys)-(p(Link%).x, p(Link%).Y), eClr
            End If

            p(Link%).Status = p(Link%).Status + 1

            tmpMaxStat% = MaxStat%
            'make bigger blast if hit city
            If p(Link%).ym > 0 Then 'bomb
                If p(Link%).Y >= p(Link%).ye Then 'hit city
                    tmpMaxStat% = MaxStat% * 2
                End If
            End If
            
            
            If p(Link%).Status <= tmpMaxStat% Then
                FrmDD.Pic.Circle (p(Link%).x, p(Link%).Y), p(Link%).Status, Clr&
            Else
                GoSub RemoveIt 'remove from list
            End If

        End If

        pLink% = Link%
        If Link% <> 0 Then
            If Link% = p(Link%).Link Then
                FrmDD.Pic.Cls
                FrmDD.Pic.Print "PROGRAM ERROR - Cross Linked List Entry Detected."
                Stop
            Else
                Link% = p(Link%).Link
            End If
        End If

    Loop
Exit Sub

RemoveIt:
    t% = p(Link%).Link
    p(Link%).Link = Pool%
    Pool% = Link%
    Nullp p(Pool%)
    If pLink% <> 0 Then
        p(pLink%).Link = t%
        Link% = pLink%
    Else
        Head% = t%
        Link% = Head%
    End If
Return

End Sub




Sub DoIt()

Dim strInfo  As String
Dim l%, bDroped%, mFired%, i%
Dim bHead%, mHead%
Dim bPool%, mPool%
Dim blnLost As Boolean, blnWon As Boolean
Dim Sync As Single
Dim x%, Y%, blnFire As Boolean
Dim CursorOff%
Dim SyncFactor As Single

    Randomize Timer
    'Randomize 7654321

    GetINI

    MachineSpeed SyncFactor

    ReDim Targets(1 To MaxTarget) As Boolean
    
    bClr = vbYellow  '
    mClr = vbCyan    '
    lClr = vbRed     'line
    eClr = vbBlack

    For i% = 1 To UBound(Targets)
        Targets(i%) = True
    Next i
    
    InitLevels Level(), SyncFactor
    
    blnAllowResize = True
    
    '................................................
    For l% = 1 To MaxLevel
    
        FrmDD.Caption = "Missile Command - Level " & l%
    
        FrmDD.Pic.BackColor = eClr
        strInfo = "Total Bombs" & vbTab & Level(l%).bMax & vbCrLf & _
            "Max Bombs" & vbTab & Level(l%).bDrop & vbCrLf & _
            "Bomb Speed" & vbTab & Level(l%).bSpeed & vbCrLf & _
            "Smart Bombs" & vbTab & Level(l%).Smart & "%" & vbCrLf & _
            "Split Bombs" & vbTab & Level(l%).Split & "%" & vbCrLf & _
            "Total Missles" & vbTab & Level(l%).mMax & vbCrLf & _
            "Max Missles" & vbTab & Level(l%).mFire & vbCrLf
        'MsgBox strInfo, vbOKOnly, "Level " & l%

        bDroped% = 0: mFired% = 0

        FrmDD.Label1(0) = Level(l%).bMax - bDroped%
        FrmDD.Label1(1) = Level(l%).mMax - mFired%

        BuildPool bPool%, bHead%, b(), Level(l%).bDrop
        BuildPool mPool%, mHead%, m(), Level(l%).mFire

        ResetTargets Targets(), gstrPath & tFile, FrmDD.Pic

        Alarm

        blnLost = False: blnWon = False

        Sync = Timer

        Do
        
            Do While Timer < Sync
                DoEvents
                If blnQuit Then Exit Do
            Loop
            If blnQuit Then Exit Do
            
            Sync = Timer + SyncDelay

            MouseRead x%, Y%, blnFire, blnQuit

        If blnLost Or blnWon Or blnQuit Then Exit Do

            LaunchB bPool%, bHead%, b(), Level(l%), bDroped%, Targets()

            If mPool% <> 0 And CursorOff% Then
                    CursorOff% = False
            ElseIf mPool% = 0 And Not CursorOff% Then
                    CursorOff% = True
            End If

            If blnFire Then
                LaunchM x%, Y%, mPool%, mHead%, m(), Level(l%), mFired%
            End If

            MyShow bHead%, bPool%, b(), bClr, bMaxStatus, mFired%, Level(l%)
            MyShow mHead%, mPool%, m(), mClr, mMaxStatus, bDroped%, Level(l%)

            Intercept mHead%, m(), bHead%, b(), Level(l%)
            If bExplodeb Then Intercept bHead%, b(), bHead%, b(), Level(l%)

            blnLost = True
            For i% = 1 To UBound(Targets)
                If Targets(i%) Then
                    blnLost = False
                    blnWon = (bDroped% = Level(l%).bMax) And (bHead% = 0)
                End If
            Next i%

        Loop

        If blnLost Or blnQuit Then Exit For
    
    Next l%
    
    If blnLost Then
        If SoundFX Then sndPlaySound gstrPath & "OhNo.wav", 1
        'MsgBox "Loser!", vbOKOnly + vbInformation, "Game Over"
    ElseIf blnQuit Then
        If SoundFX Then sndPlaySound gstrPath & "Error.wav", 1
        'MsgBox "Quiter!", vbOKOnly + vbInformation, "Chicken!"
    ElseIf blnWon Then
        'If SoundFX Then sndPlaySound gstrPath & "OnNo.wav", 1
        'MsgBox "Cheater!", vbOKOnly + vbInformation, "You Won?"
    End If
    
    'other features:
    '   scoring
    '   panic button (shoot every thing at specified altitude in pattern)
    '   Attacks within levels
    '   multiple ammo dumps near cities that can be destroyed and use up
    '       players ammo
    '   smart missles that track bombs (limited number?)
    '   super missles that have big explosions (limited number?)

End Sub


Function VBGetPrivateProfileString(Section$, key$, File$) As String
    
    Dim KeyValue$
    Dim characters As Long
    
    KeyValue$ = String$(128, 0)
    
    characters = GetPrivateProfileStringByKeyName(Section$, key$, "", KeyValue$, 127, File$)

    If characters > 0 Then
        KeyValue$ = Left$(KeyValue$, characters)
    End If
    
    VBGetPrivateProfileString = KeyValue$

End Function

Sub GetINI(Optional blnReset As Boolean = False)

Dim i As Integer
Dim Section As String
Dim SectionList As String
Dim INIFile As String
Dim value
    
'/////////////////
l(1) = "10, 50, 5, 5, 0.05,  0, 15, ""Slow and Dumb I"""
l(2) = "10, 50, 5, 10, 0.05,  0, 20, ""Slow and Dumb II"""
l(3) = "10, 50, 5, 10, 0.07,  30, 25, ""Faster and Smarter I"""
l(4) = "15, 100, 10, 20, 0.09,  40, 25, ""Faster and Smarter II"""
l(5) = "15, 100, 10, 20, 0.09,  50, 25, ""Faster and Smarter III"""
l(6) = "20, 200, 10, 20, 0.12,  50, 33, ""Prelude"""
l(7) = "30, 200, 10, 20, 0.12,  60, 33, ""Dooms Day I"""
l(8) = "40, 200, 20, 30, 0.12,  70, 33, ""Dooms Day II"""
l(9) = "50, 200, 50, 50, 0.12,  80, 33, ""You've got to be kidding!"""
l(10) = "100, 200, 50, 50, 0.15,  90, 50, ""This ain't right"""
SoundFX = True
MaxTarget = 10
SyncDelay = 0.05
SyncDist = 500
SyncTime = 1
bMaxStatus = 35
mMaxStatus = 25
mSpeed = 1.5
bExplodeb = True
'/////////////////
    
    INIFile = gstrPath & "DD.ini"

    If Not blnReset And Exist(INIFile) Then
    
        Section = "DD"
    
        SectionList = String(6144, 0)
    
        Call GetPrivateProfileStringSections(0, 0, "", SectionList, 6144, INIFile)

        If InStr(SectionList, Section) > 0 Then
            For i = 1 To MaxLevel
                value = VBGetPrivateProfileString(Section, "l" & i, INIFile)
                If Asc(value) <> 0 Then
                    l(i) = value
                End If
            Next i
        End If

        value = VBGetPrivateProfileString(Section, "Sound", INIFile)
        If Asc(value) <> 0 Then
            SoundFX = value
        End If

        value = VBGetPrivateProfileString(Section, "Cities", INIFile)
        If Asc(value) <> 0 Then
            If value >= 5 And value <= 20 Then
                MaxTarget = value
            End If
        End If
        
        value = VBGetPrivateProfileString(Section, "SyncDelay", INIFile)
        If Asc(value) <> 0 Then
            SyncDelay = value
        End If
        
        value = VBGetPrivateProfileString(Section, "SyncDist", INIFile)
        If Asc(value) <> 0 Then
            SyncDist = value
        End If
        
        value = VBGetPrivateProfileString(Section, "SyncTime", INIFile)
        If Asc(value) <> 0 Then
            SyncTime = value
        End If

        value = VBGetPrivateProfileString(Section, "bRadius", INIFile)
        If Asc(value) <> 0 Then
            If value >= 5 And value <= 50 Then
                bMaxStatus = value
            End If
        End If

        value = VBGetPrivateProfileString(Section, "mRadius", INIFile)
        If Asc(value) <> 0 Then
            If value >= 5 And value <= 50 Then
                mMaxStatus = value
            End If
        End If

        value = VBGetPrivateProfileString(Section, "mSpeed", INIFile)
        If Asc(value) <> 0 Then
            If value >= 0.1 And value <= 5 Then
                mSpeed = value
            End If
        End If
        
        value = VBGetPrivateProfileString(Section, "bExplodeb", INIFile)
        If Asc(value) <> 0 Then
            bExplodeb = value
        End If

    Else
        PutINI
    End If


End Sub
Sub PutINI()

Dim Section As String
Dim INIFile As String
Dim Tmp As Long
Dim i As Integer
    
    INIFile = gstrPath & "DD.ini"
    Section = "DD"

    Tmp = WritePrivateProfileStringByKeyName(Section, "REM", "rem=bMax, mMax, bDrop, mFire, bSpeed, Smart%, Split%, Name", INIFile)

    For i = 1 To MaxLevel
        Tmp = WritePrivateProfileStringByKeyName(Section, "l" & i, l(i), INIFile)
    Next i
    
    Tmp = WritePrivateProfileStringByKeyName(Section, "Sound", SoundFX, INIFile)
    Tmp = WritePrivateProfileStringByKeyName(Section, "Cities", MaxTarget, INIFile)
    
    Tmp = WritePrivateProfileStringByKeyName(Section, "SyncDist", SyncDist, INIFile)
    Tmp = WritePrivateProfileStringByKeyName(Section, "SyncDelay", SyncDelay, INIFile)
    Tmp = WritePrivateProfileStringByKeyName(Section, "SyncTime", SyncTime, INIFile)
    
    Tmp = WritePrivateProfileStringByKeyName(Section, "mRadius", mMaxStatus, INIFile)
    Tmp = WritePrivateProfileStringByKeyName(Section, "bRadius", bMaxStatus, INIFile)
    
    Tmp = WritePrivateProfileStringByKeyName(Section, "mSpeed", mSpeed, INIFile)
    
    Tmp = WritePrivateProfileStringByKeyName(Section, "bExplodeb", bExplodeb, INIFile)

End Sub
Public Function Exist(f As String) As Boolean


    Exist = False
    On Error Resume Next
    Exist = Len(Dir(f)) <> 0

End Function

