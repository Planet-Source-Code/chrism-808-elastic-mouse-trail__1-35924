VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   576
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSlide 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   320
      Left            =   60
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   3
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   60
      TabIndex        =   1
      Text            =   "your message here"
      Top             =   300
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Elasticity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'letterfollow type holds the stuff for each letter
Private Type LetterFollow
    X As Single
    Y As Single
    Letter As String
End Type

'array of 101 letters except -1 is for the mouse position
Dim Letter(-1 To 100) As LetterFollow
'how elastic the string is
Dim Elasticity As Integer
'to know when to finish
Dim Finish As Boolean
'stuff to do withe my slider bar
Dim SlideChange As Boolean
Dim SlidePos As Integer


Private Sub Form_Activate()
Dim i As Integer
Dim t As Integer

Do
    'windows does stuff
    DoEvents
    'clear old writing
    Form1.Cls
    
    'if the message is longer than 100 cut it to 100
    If Len(Text1.Text) > 100 Then
        Text1.Text = Mid(Text1.Text, 1, 100)
    End If
    
    'loop through every letter
    For i = 0 To Len(Text1.Text)
        'give each letter its letter
        Letter(i).Letter = Mid(Text1.Text, i + 1, 1)

        'this calculates the x position by taking its current
        'position + the difference between where it is and
        'where it wants to go divided by the elasticity
        'the 8 is x distance between letters
        Letter(i).X = Letter(i).X + (Letter(i - 1).X + 8 - Letter(i).X) / Elasticity
        'same as no 8 since 0 is x distance between letters
        Letter(i).Y = Letter(i).Y + (Letter(i - 1).Y - Letter(i).Y) / Elasticity
        
        'sets the current position to the letters x and y
        Form1.CurrentX = Letter(i).X
        Form1.CurrentY = Letter(i).Y
        'prints the letter
        Form1.Print Letter(i).Letter
    Next i
    t = t + 1
    If t = 60 Then Finish = True
Loop Until Finish = True
'close form since person clicked close
'Unload Me
End Sub

Private Sub Form_Load()
'sets sliderbar position to 1
SlideBar picSlide, 10, 20
'sets elasticity
Elasticity = 2
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'when mouse moves get the new x and y pos and put into
'the -1 position
Letter(-1).X = X
Letter(-1).Y = Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
'unloading form so stop loop
Finish = True
End Sub

Private Sub picSlide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If picSlide.Point(X, Y) = vbBlack Then
    'user has clicked down on the slider
    SlideChange = True
End If
End Sub

Private Sub picSlide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim BuffX As Integer

If SlideChange = True Then
    'for some reason when i just use x it does not work
    'so i created the variable buffx
    BuffX = X
    'sets the limit on it
    If BuffX > picSlide.ScaleWidth Then
        BuffX = picSlide.ScaleWidth - 1
    ElseIf BuffX < 0 Then
        BuffX = 0
    End If
    
    'calls the function that does everything
    SlideBar picSlide, BuffX, 20
    
End If
End Sub

Private Sub picSlide_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'mouse up so not moving the slider
SlideChange = False
End Sub

Private Function SlideBar(Slide As PictureBox, ScrollPos As Integer, NumSnap As Integer)
Dim SnapDis As Single
Dim i As Single
Dim Distance As Integer
Dim SlideHeight As Integer

'clear old ticks
Slide.Cls

'so its 1 and up that will have a snap
If NumSnap > 0 Then
    'calculates the distasnce between snaps
    SnapDis = Slide.ScaleWidth / (NumSnap + 1)

    'finds the closest snapPoint to the scrollpos
    For i = 0 To Slide.ScaleWidth + 1 Step SnapDis
        'draws the little tick lines
        Slide.Line (i, Slide.ScaleHeight - 3)-(i, Slide.ScaleHeight), RGB(0, 1, 0)
        
        'calculate distance from check to x
        Distance = Sqr((i - ScrollPos) ^ 2)
        'if less than half a check then this is the
        'right check
        If Distance <= SnapDis / 2 Then
            ScrollPos = i
            Elasticity = (i / SnapDis) + 1
            SlidePos = ScrollPos
        End If
    Next i
    
    'draws last check
    Slide.Line (Slide.ScaleWidth - 1, Slide.ScaleHeight - 2)-(Slide.ScaleWidth - 1, Slide.ScaleHeight), RGB(0, 1, 0)

End If

If ScrollPos > Slide.ScaleWidth - 2 Then
    ScrollPos = Slide.ScaleWidth - 2
ElseIf ScrollPos < 2 Then
    ScrollPos = 2
End If

SlidePos = ScrollPos

SlideHeight = Slide.ScaleHeight - 6

'first part of the slide bar
Slide.Line (1, 1)-(ScrollPos - 1, SlideHeight), vbBlue, BF
Slide.Line (1, 1)-(ScrollPos - 3, SlideHeight), &HFF8080, B

'second part of the slide bar (after move line)
Slide.Line (ScrollPos + 1, 1)-(Slide.ScaleWidth - 1, SlideHeight), vbBlue, BF
Slide.Line (ScrollPos + 2, 1)-(Slide.ScaleWidth - 2, SlideHeight), &HFF8080, B

'black line
Slide.Line (0, 0)-(Slide.ScaleWidth - 1, SlideHeight + 1), RGB(0, 1, 0), B

'black scrolling line
Slide.Line (ScrollPos - 2, 1)-(ScrollPos + 1, SlideHeight), vbBlack, BF

End Function

