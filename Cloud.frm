VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Click on picture to save image"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3720
      Top             =   6360
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      Height          =   5880
      Left            =   120
      ScaleHeight     =   5820
      ScaleMode       =   0  'User
      ScaleWidth      =   8533.333
      TabIndex        =   0
      Top             =   120
      Width           =   7200
   End
   Begin VB.Label Label1 
      Caption         =   "Wait..."
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   6360
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Fractal - DUST CLOUD by Dolac

'picture is derived from single user defined starting point P0(A,B,X,Y).
'by numerous repeating of some transformations we get some cool pics.
'Because of the fact that we can follow the trail start point is leaving
'behined (P0, P1, P2, .... Px - dots on picture) this pics are also called
'orbitals of dynamic system

Option Explicit
Dim Xs As Double:   Dim Ys As Double    'picturebox start point

Dim N As Double     'number of cycles
Dim K As Single     'number of iterative steps in one cycle

Dim A As Double     'user adjustable parametar
Dim B As Double     'user adjustable parametar
Dim X As Double     'user adjustable parametar
Dim Y As Double     'user adjustable parametar

Dim W As Double     'transformation variable - some function
Dim Z As Double     'hold previous step of X

Dim ScalePic As Integer 'scale graphics in picturebox - rude zoom
Dim Pattern As Integer  'try patterns 1 to 6 and you will get it

Private Sub Form_Load()
    Timer1.Enabled = True:      Timer1.Interval = 1
    Xs = 4200:                  Ys = 2900:           N = 0:
        
    Call NextOrbit
    
    Pattern = InputBox("Selet Pattern 1 - 6", 100, 4, 300)
        
        'to get a rude estimation what A, B, X and Y are good to use - I prepared you Clouds.xls
        'sheets 1 to 6 shows computation sequence series of here given pattern examples 1 to 6
        'X() sheets show graphic look of computation sequence of poorly picked parameters
        
        Select Case Pattern
            Case 1
                A = 3:      B = -2.73:   X = 3.21:      Y = 6.54:       ScalePic = 250:    K = 44
            Case 2
                A = 3.08:   B = -3.96:   X = 12.61:     Y = 18.54:      ScalePic = 13:     K = 111
            Case 3.1
                A = 2.3:    B = -2.9:    X = 6.21:      Y = 14.54:      ScalePic = 62:     K = 111
            Case 4
                A = 3.7:    B = -3.5:    X = 1:         Y = 22:         ScalePic = 66:     K = 55
            Case 5
                A = 0.8:    B = -1:      X = 1:         Y = 2:          ScalePic = 1000:   K = 69
            Case 6
                A = 1.2:    B = -2.1:    X = 33:        Y = 3.1:        ScalePic = 19:      K = 277
            
            'Try some of your own
            'Case 7
            '    A = 1.2:    B = -2.1:    X = 33:        Y = 11:         ScalePic = 42:      K = 54
        End Select
End Sub

Private Function Draw()
    Dim i As Single
    
    For i = 1 To 2000       'used to speed up making of drawing, and also make computer accessible
        
        'draw dot on Picture1
        Picture1.PSet (Xs + X * ScalePic, Ys + Y * ScalePic), vbBlack:       Picture1.Refresh
        
            Z = X:      X = Y + W:      Call NextOrbit
        
            Y = W - Z
        
        'Debug.Print Format(A, "#,###") & "   " & Format(B, "#.00") & "  " & Format(X, "#.00") _
                      & "   " & Format(Y, "#.00") & "    " & Format(W, "#.00")
    Next i
End Function

Private Function NextOrbit()
    If X > 1 Then
        W = A * X + B * (X - 1)     'can u even imagine how many other/different functions can we use here...
    End If
    
    If X < -1 Then
        W = A * X + B * (X + 1)     'maybe some combination involving sin or cos...
    End If
    
    If X < 1 And X > -1 Then
        W = A * X                   'or ln/log or exp... just to give you a hint..
    End If
End Function

Private Sub Timer1_Timer()
    Call Draw
    
    N = N + 1:    Label1.Caption = "Wait - Iteration cycle " & N & " / " & K
    
    If N = K Then       'stop condition
        Timer1.Enabled = False
        Label1.Caption = "Iteration ended"
    End If
End Sub

Private Sub Picture1_Click()
    Dim PicFile As String
    'save bmp of picture
    PicFile = App.Path & "\Pics\Pattern-" & Pattern & "  Scale-" & ScalePic & " Cycles-" & K & ".bmp"
    SavePicture Picture1.Image, PicFile
    MsgBox "Picture saved.."
End Sub

Private Sub Form_Click()
    'Debug.Print vbCrLf & "  N     X       W      Y" & vbCrLf
    Unload Me
End Sub
