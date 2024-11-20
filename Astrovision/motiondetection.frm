VERSION 5.00
Object = "{AB6BEDFD-8860-4D37-BBA5-EDC1EE715F32}#1.0#0"; "gridbots Webcam API.ocx"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   10065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   LinkTopic       =   "Form4"
   ScaleHeight     =   10065
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   2760
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Left            =   9120
      Top             =   8040
   End
   Begin VB.Timer Timer1 
      Left            =   9120
      Top             =   7440
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Detection mode"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   6
      Top             =   8760
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00E0E0E0&
      Height          =   3375
      Left            =   2760
      ScaleHeight     =   221
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   221
      TabIndex        =   5
      Top             =   240
      Width           =   3375
   End
   Begin VB.PictureBox Picture3 
      Height          =   735
      Left            =   3720
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Loaded "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   3
      Top             =   7920
      Width           =   1575
   End
   Begin gridBotsWebcamAPI.gridBotsWebCam gridBotsWebCam1 
      Height          =   5445
      Left            =   360
      TabIndex        =   2
      Top             =   4080
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   9604
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   3375
      Left            =   6240
      ScaleHeight     =   221
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   221
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   3375
      Left            =   6240
      ScaleHeight     =   221
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   221
      TabIndex        =   0
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      Height          =   1695
      Left            =   6840
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      Height          =   6015
      Left            =   120
      Top             =   3840
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Motion Detected !!!!"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   2415
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim g(4000, 4000) As Long
Dim g1(4000, 4000) As Long
Dim g2(4000, 4000) As Long
Dim g3(4000, 4000) As Long
Dim val As Long, val2 As Long, rrr As Long, ggg As Long, bbb As Long
Dim grayscale As Long, grayscale2 As Long
Dim c As Long
Dim X As Long, a As Long
Dim i As Long, j As Long
Dim imax As Long, imin As Long, jmax As Long, jmin As Long
Private Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Sub Command4_Click()
Unload Form4
Load Form1
Form1.Show
End Sub

Private Sub Form_Load()
Picture3.Visible = False
Form4.BackColor = RGB(70, 150, 240)
Label1.Visible = False
End Sub
Private Sub Command1_Click()
'Timer1.Enabled = True
'Timer1.Interval = 5000
'Timer2.Enabled = True
'Timer2.Interval = 8000
Label1.Visible = False
Picture3.Picture = gridBotsWebCam1.getPIC
Picture2.Picture = Picture1.Picture
Picture1.Picture = Picture3.Picture
End Sub
Private Sub Command2_Click()
'MsgBox (" Detection mode Activated")
Label1.Visible = False
c = 0
For i = 0 To Picture1.Picture.Height / 30
For j = 0 To Picture1.Picture.Width / 30
val = GetPixel(Picture1.hDC, i, j)
val2 = GetPixel(Picture2.hDC, i, j)
rr = val And 255
gg = (val \ 256) And 255
bb = (val \ 65536) And 255

rrr = val2 And 255
ggg = (val2 \ 256) And 255
bbb = (val2 \ 65536) And 255

grayscale = (rr / 3 + gg / 3 + bb / 3)
grayscale2 = (rrr / 3 + ggg / 3 + bbb / 3)
g(i, j) = grayscale
g1(i, j) = grayscale2
Picture1.PSet (i, j), RGB(grayscale, grayscale, grayscale) ' gray scale pic 1
Picture2.PSet (i, j), RGB(grayscale2, grayscale2, grayscale2)  ' gray scale pic 2
Next j
Next i
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
' get frame2-frame1

Picture4.Picture = Picture1.Picture
For i = 0 To Picture3.Picture.Height / 30
For j = 0 To Picture3.Picture.Width / 30
  
  If g1(i, j) - g(i, j) > 0 Then
        g2(i, j) = g1(i, j) - g(i, j)     'g2(i, j) = motion
   
        If g2(i, j) > 50 Then
            c = c + 1
            g3(i, j) = 255
            Picture4.PSet (i, j), vbRed
        End If
  Else
        grayscale = 0
        g2(i, j) = grayscale
        g3(i, j) = 0                      'why ???? forgot
  End If
Next j
Next i

Text1.Text = c
If c > 100 Then
    Label1.Visible = True
End If

imin = 1000
jmin = 1000
imax = 0
jmax = 0
    For i = 1 To Picture1.Picture.Width / 20
    For j = 1 To Picture1.Picture.Height / 20
If g3(i, j) > 0 Then
        If g3(i, j) > 0 Then
            If i > imax Then
            imax = i
            End If
        End If
        If g3(i, j) > 0 Then
            If j > jmax Then
            jmax = j
            End If
        End If
        If g3(i, j) > 0 Then
            If j < jmin Then
            jmin = j
            End If
            End If
        If g3(i, j) > 0 Then
            If i < imin Then
            imin = i
            End If
        End If
Else
'Picture2.PSet (i, j), vbBlack   ' for rest of picture PIC-BLOB
End If
    Next j
    Next i
Picture4.Line (imin, jmin)-(imax, jmin), vbRed
Picture4.Line (imax, jmin)-(imax, jmax), vbRed
Picture4.Line (imax, jmax)-(imin, jmax), vbRed
Picture4.Line (imin, jmax)-(imin, jmin), vbRed
End Sub
Private Sub Form4_Terminate()
End
End Sub
