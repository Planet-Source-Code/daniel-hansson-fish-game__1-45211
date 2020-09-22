Attribute VB_Name = "Stars"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long
    Const PI = 3.14159          'Mmmm.. Pi
    Const MS_DELAY = 20         'Milliseconds per frame (25 = 40 frames per second)
    Const STAR_RADIUS = 0.5       'Radius of the stars
    Const NUM_STARS = 40       'Number of stars in the field
    Const NUM_LAYERS = 3        'Number of speed layers in the field
    
    Dim mlngTimer As Long       'Holds system time since last frame was displayed
    Dim msngHeading As Single   'Current direction in which ship is moving
    Dim msngSpeed As Single     'Current speed with which ship is moving
    Dim mblnRunning As Boolean  'Is the render loop running?
    Dim mblnUpKey As Boolean    'Is the up arrow-key depressed? ... Are YOU depressed? Fly my triangle around for a while, you'll feel better!


    Private Type STAR_TYPE
        sngX As Single          'X Coord of the star
        sngY As Single          'Y Coord of the star
        sngRelSpeed As Single   'Speed of the star relative to the speed of the ship
    End Type
    Dim mudtStar() As STAR_TYPE 'An array of stars!

Private Sub Physics()

Dim sngXComp As Single  'Resultant X and Y components
Dim sngYComp As Single
Dim i As Integer
    'Thrust
    If mblnUpKey Then
        'Determine the X and Y components of the resultant vector
        'sngXComp = msngSpeed * Sin(msngHeading) + ACCEL * Sin(msngFacing)
        sngXComp = Sin(1) * Sin(PI / 360 * 180) * 4
        'sngYComp = msngSpeed * Cos(msngHeading) + ACCEL * Cos(msngFacing)
        sngYComp = Cos(1) * Cos(PI / 360 * 180) * 4
        'Determine the resultant speed
        msngSpeed = Sqr(sngXComp ^ 2 + sngYComp ^ 2)
        'Calculate the resultant heading, and adjust for arctangent by adding Pi if necessary
        If sngYComp > 0 Then msngHeading = Atn(sngXComp / sngYComp)
        If sngYComp < 0 Then msngHeading = Atn(sngXComp / sngYComp) + PI
    End If
    
    'Move the stars
    For i = 0 To UBound(mudtStar)
        'Move the stars according to their relative speeds (w.r.t. the inverse of the ship's speed)
        mudtStar(i).sngX = mudtStar(i).sngX - msngSpeed * mudtStar(i).sngRelSpeed * Sin(msngHeading)
        mudtStar(i).sngY = mudtStar(i).sngY + msngSpeed * mudtStar(i).sngRelSpeed * Cos(msngHeading)
        'Wrap the stars at the edges of the window
        If mudtStar(i).sngX > frmMain.PicDisp.ScaleWidth Then mudtStar(i).sngX = 0
        If mudtStar(i).sngY > frmMain.PicDisp.ScaleHeight Then mudtStar(i).sngY = 0
        If mudtStar(i).sngX < 0 Then mudtStar(i).sngX = frmMain.PicDisp.ScaleWidth
        If mudtStar(i).sngY < 0 Then mudtStar(i).sngY = frmMain.PicDisp.ScaleHeight
    Next i
    
End Sub


Private Sub DrawStars()

Dim i As Integer

    'Display every star in the array
    For i = 0 To UBound(mudtStar)
        frmMain.PicDisp.Circle (mudtStar(i).sngX, mudtStar(i).sngY), STAR_RADIUS, RGB(40, 20, 10)
    Next i
    
End Sub

Public Sub ReDrawStars()
If mlngTimer + MS_DELAY <= GetTickCount() Then
    mlngTimer = GetTickCount()  'Reset the timer variable
    Physics                     'Allow the ship's location to be updated
    DrawStars                   'Draw the starfield
End If
End Sub


Public Sub SetStars()
    mblnUpKey = True
    Dim i As Integer
    mlngTimer = GetTickCount()
    Randomize
    ReDim mudtStar(NUM_STARS - 1)
    For i = 0 To UBound(mudtStar)
        'Set the star's random X and Y coords
        mudtStar(i).sngX = Rnd() * frmMain.PicDisp.ScaleWidth
        mudtStar(i).sngY = Rnd() * frmMain.PicDisp.ScaleHeight
        'Set the star's relative speed
        mudtStar(i).sngRelSpeed = ((i \ (NUM_STARS \ NUM_LAYERS)) + 1) / NUM_LAYERS
    Next i
End Sub
