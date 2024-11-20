VERSION 5.00
Object = "{AB6BEDFD-8860-4D37-BBA5-EDC1EE715F32}#1.0#0"; "gridbots Webcam API.ocx"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11325
   LinkTopic       =   "Form3"
   ScaleHeight     =   9810
   ScaleWidth      =   11325
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
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Left            =   6480
      Top             =   1680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GET"
      Height          =   855
      Left            =   6000
      TabIndex        =   3
      Top             =   4560
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Height          =   3735
      Left            =   4200
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   2
      Top             =   5640
      Width           =   3735
   End
   Begin gridBotsWebcamAPI.gridBotsWebCam gridBotsWebCam1 
      Height          =   5445
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   9604
   End
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   360
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   0
      Top             =   5640
      Width           =   3735
   End
   Begin VB.Shape Shape4 
      Height          =   375
      Left            =   7080
      Top             =   480
      Width           =   495
   End
   Begin VB.Shape Shape3 
      Height          =   375
      Left            =   5880
      Top             =   480
      Width           =   495
   End
   Begin VB.Shape Shape2 
      Height          =   375
      Left            =   8280
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "min"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   8
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Hrs"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   480
      Width           =   495
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   5760
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8400
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   5
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim val As Long, rr As Long, gg As Long, bb As Long, img(4000, 4000) As Long, grayscale As Long, transpose(4000, 4000) As Long
Dim a(4000, 4000) As Long, c(4000, 4000) As Long, leftr(4000, 4000) As Long
Private Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Sub Command1_Click()
Picture1.Picture = gridBotsWebCam1.getPIC

For j = 0 To Picture1.Picture.Width / 20
For i = 0 To Picture1.Picture.Height / 20
val = GetPixel(Picture1.hDC, i, j)
rr = val And 255
gg = (val \ 256) And 255
bb = (val \ 65536) And 255
grayscale = (rr / 3 + gg / 3 + bb / 3)
img(i, j) = grayscale
a(i, i) = 1
Picture1.PSet (i, j), RGB(grayscale, grayscale, grayscale)
Next i
Next j
' Transpose
For j = 0 To Picture1.Picture.Width / 20
For i = 0 To Picture1.Picture.Height / 20
transpose(i, j) = img(j, i)
Picture2.PSet (i, j), RGB(transpose(i, j), transpose(i, j), transpose(i, j))
Next i
Next j
End Sub

Private Sub Command4_Click()
Unload Form3
Load Form1
Form1.Show
End Sub

Private Sub Form_Load()
Form3.BackColor = RGB(70, 150, 240)
Label1.Caption = Hour(Now)
Label2.Caption = Minute(Now)
Label3.Caption = DateValue(Now)
Timer1.Enabled = True
Timer1.Interval = 1000
End Sub

Private Sub Form3_Terminate()
End
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Hour(Now)
Label2.Caption = Minute(Now)
End Sub
