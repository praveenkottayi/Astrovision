VERSION 5.00
Object = "{AB6BEDFD-8860-4D37-BBA5-EDC1EE715F32}#1.0#0"; "gridbots Webcam API.ocx"
Begin VB.Form Form2 
   Caption         =   "LIVE"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   ScaleHeight     =   734
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      TabIndex        =   16
      Top             =   120
      Width           =   2175
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000FF&
      Height          =   4575
      Left            =   10680
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   13
      Top             =   5760
      Width           =   3975
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   4455
      Left            =   10560
      ScaleHeight     =   293
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   12
      Top             =   600
      Width           =   3975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "THRESHOLD"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   10
      Top             =   9720
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Smoothing"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   9
      Top             =   8280
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C00000&
      Caption         =   "Greyscale"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      MaskColor       =   &H00FF0000&
      TabIndex        =   8
      Top             =   8280
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   4200
      TabIndex        =   7
      Top             =   7080
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   2640
      TabIndex        =   6
      Top             =   7080
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   1080
      TabIndex        =   5
      Top             =   7080
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   4
      Text            =   "200"
      Top             =   9600
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      Height          =   3255
      Left            =   6000
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   253
      TabIndex        =   3
      Top             =   7560
      Width           =   3855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   3255
      Left            =   6000
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   253
      TabIndex        =   2
      Top             =   4080
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   3255
      Left            =   6000
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   253
      TabIndex        =   1
      Top             =   480
      Width           =   3855
   End
   Begin gridBotsWebcamAPI.gridBotsWebCam gridBotsWebCam1 
      Height          =   5445
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   9604
   End
   Begin VB.Line Line1 
      X1              =   224
      X2              =   400
      Y1              =   680
      Y2              =   712
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "SMOOTHENED"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6000
      TabIndex        =   18
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "GREYSCALE"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6000
      TabIndex        =   17
      Top             =   240
      Width           =   1935
   End
   Begin VB.Shape Shape10 
      Height          =   10815
      Left            =   5760
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Histogram after SMOOTHING "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   10560
      TabIndex        =   15
      Top             =   5280
      Width           =   4455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Histogram of GRAYSCALE image"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   9840
      TabIndex        =   14
      Top             =   120
      Width           =   5415
   End
   Begin VB.Shape Shape9 
      Height          =   5175
      Left            =   10320
      Shape           =   4  'Rounded Rectangle
      Top             =   5280
      Width           =   4455
   End
   Begin VB.Shape Shape8 
      Height          =   5175
      Left            =   10320
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4455
   End
   Begin VB.Shape Shape6 
      Height          =   855
      Left            =   480
      Top             =   8160
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Enter threshold value :10 -255"
      BeginProperty Font 
         Name            =   "WST_Swed"
         Size            =   9
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   9360
      Width           =   4935
   End
   Begin VB.Shape Shape5 
      Height          =   1335
      Left            =   480
      Top             =   9120
      Width           =   5055
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FF0000&
      Height          =   975
      Left            =   4080
      Top             =   6960
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C000&
      Height          =   975
      Left            =   2520
      Top             =   6960
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      Height          =   975
      Left            =   960
      Top             =   6960
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   480
      Top             =   6840
      Width           =   5055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Clr As Long, r As Long, g As Long, b As Long, gry As Long
Dim s As Long, h(256) As Long, rr As Long, gg As Long, bb As Long, grayscale As Long, val As Long
Dim img(4000, 4000), limit As Long
Private Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Sub Command4_Click()
Unload Form2
Load Form1
Form1.Show
End Sub

Private Sub Form2_Terminate()
End
End Sub

Private Sub gridBotsWebCam1_pictureCaptured()
Picture1.Picture = gridBotsWebCam1.getPIC
End Sub
Private Sub Form_Load()
'Path = "D:\Astrovision\Images\8.jpg"
'Path = "C:\Documents and Settings\Administrator\Desktop\Astrovision\Images\8.jpg"
'Form2.Picture = LoadPicture(Path)
Form2.BackColor = RGB(70, 150, 240)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Clr = GetPixel(Form2.hDC, X, Y)
r = Clr And 255
g = (Clr \ 256) And 255
b = (Clr \ 65536) And 255
gry = (r / 3 + g / 3 + b / 3)
Text2.BackColor = (RGB(r, 0, 0))
Text3.BackColor = (RGB(0, g, 0))
Text4.BackColor = (RGB(0, 0, b))
End Sub
Private Sub Command1_Click()
' Initialize h()
For s = 0 To 255
h(s) = 0
Next s
Picture4.Refresh
' Greyscale and Histogram
For i = 0 To Picture1.Picture.Height / 50
For j = 0 To Picture1.Picture.Width / 50
val = GetPixel(Picture1.hDC, i, j)
rr = val And 255
gg = (val \ 256) And 255
bb = (val \ 65536) And 255
grayscale = (rr / 3 + gg / 3 + bb / 3)
img(i, j) = grayscale
Picture1.PSet (i, j), RGB(grayscale, grayscale, grayscale)
h(grayscale) = h(grayscale) + 1
Picture4.Line (grayscale, 0)-(grayscale, h(grayscale))
Next j
Next i
MsgBox ("Histogram completed")
Form2.Command2.Enabled = True
End Sub
Private Sub Command2_Click()
For s = 0 To 255
h(s) = 0
Next s
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'smoothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Picture5.Refresh
For i = 1 To Picture1.Picture.Height / 50
For j = 1 To Picture1.Picture.Width / 50
If img(i, j) > 0 Then
img(i - 1, j) = 0.1 * (img(i, j) + img(i, j + 1) + img(i, j + 2) + img(i + 1, j) + img(i + 1, j + 1) + img(i + 1, j + 2) + img(i + 2, j) + img(i + 2, j + 1) + img(i + 2, j + 2)) / 9
img(i - 1, j - 1) = 0.08 * (img(i, j) + img(i, j + 1) + img(i, j + 2) + img(i + 1, j) + img(i + 1, j + 1) + img(i + 1, j + 2) + img(i + 2, j) + img(i + 2, j + 1) + img(i + 2, j + 2)) / 9
img(i, j - 1) = 0.1 * (img(i, j) + img(i, j + 1) + img(i, j + 2) + img(i + 1, j) + img(i + 1, j + 1) + img(i + 1, j + 2) + img(i + 2, j) + img(i + 2, j + 1) + img(i + 2, j + 2)) / 9
img(i + 1, j - 1) = 0.08 * (img(i, j) + img(i, j + 1) + img(i, j + 2) + img(i + 1, j) + img(i + 1, j + 1) + img(i + 1, j + 2) + img(i + 2, j) + img(i + 2, j + 1) + img(i + 2, j + 2)) / 9
img(i + 1, j) = 0.1 * (img(i, j) + img(i, j + 1) + img(i, j + 2) + img(i + 1, j) + img(i + 1, j + 1) + img(i + 1, j + 2) + img(i + 2, j) + img(i + 2, j + 1) + img(i + 2, j + 2)) / 9
img(i + 1, j + 1) = 0.08 * (img(i, j) + img(i, j + 1) + img(i, j + 2) + img(i + 1, j) + img(i + 1, j + 1) + img(i + 1, j + 2) + img(i + 2, j) + img(i + 2, j + 1) + img(i + 2, j + 2)) / 9
img(i, j + 1) = 0.1 * (img(i, j) + img(i, j + 1) + img(i, j + 2) + img(i + 1, j) + img(i + 1, j + 1) + img(i + 1, j + 2) + img(i + 2, j) + img(i + 2, j + 1) + img(i + 2, j + 2)) / 9
img(i - 1, j + 1) = 0.08 * (img(i, j) + img(i, j + 1) + img(i, j + 2) + img(i + 1, j) + img(i + 1, j + 1) + img(i + 1, j + 2) + img(i + 2, j) + img(i + 2, j + 1) + img(i + 2, j + 2)) / 9
img(i, j) = 2.64 * (img(i, j) + img(i, j + 1) + img(i, j + 2) + img(i + 1, j) + img(i + 1, j + 1) + img(i + 1, j + 2) + img(i + 2, j) + img(i + 2, j + 1) + img(i + 2, j + 2)) / 9
Else
img(i, j) = 0
End If
Picture2.PSet (i, j), RGB(img(i, j), img(i, j), img(i, j))
grayscale = img(i, j)
If grayscale > 255 Then
    grayscale = grayscale - 255
End If
h(grayscale) = h(grayscale) + 1
Picture5.Line (grayscale, 0)-(grayscale, h(grayscale))
Next j
Next i
'''''''''''''''''''''''''''''''''''''''''''''''''''''''successful..............
MsgBox (" Smoothing and Histohram completed")
End Sub
Private Sub Command3_Click()

' for thresholding

limit = Text1.Text
If limit < 10 Or limit > 256 Then
limit = 120
MsgBox (" Invalid Data (120 taken as default)")
End If
For i = 0 To Picture1.Picture.Height / 30
For j = 0 To Picture1.Picture.Width / 30
val = GetPixel(Picture1.hDC, i, j)
rr = val And 255
gg = (val \ 256) And 255
bb = (val \ 65536) And 255
grayscale = (rr / 3 + gg / 3 + bb / 3)
If (grayscale > limit) Then
      grayscale = 255
      'g(i, j) = 255
Else
      grayscale = 0
      'g(i, j) = 0
End If
Picture3.PSet (i, j), RGB(grayscale, grayscale, grayscale)
Next j
Next i
MsgBox ("Grayscale complete !!")
End Sub



Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Clr = GetPixel(Picture1.hDC, X, Y)
r = Clr And 255
g = (Clr \ 256) And 255
b = (Clr \ 65536) And 255
gry = (r / 3 + g / 3 + b / 3)
Text2.BackColor = (RGB(r, 0, 0))
Text3.BackColor = (RGB(0, g, 0))
Text4.BackColor = (RGB(0, 0, b))
End Sub
Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Clr = GetPixel(Picture2.hDC, X, Y)
r = Clr And 255
g = (Clr \ 256) And 255
b = (Clr \ 65536) And 255
gry = (r / 3 + g / 3 + b / 3)
Text2.BackColor = (RGB(r, 0, 0))
Text3.BackColor = (RGB(0, g, 0))
Text4.BackColor = (RGB(0, 0, b))
End Sub
Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Clr = GetPixel(Picture3.hDC, X, Y)
r = Clr And 255
g = (Clr \ 256) And 255
b = (Clr \ 65536) And 255
gry = (r / 3 + g / 3 + b / 3)
Text2.BackColor = (RGB(r, 0, 0))
Text3.BackColor = (RGB(0, g, 0))
Text4.BackColor = (RGB(0, 0, b))
End Sub
