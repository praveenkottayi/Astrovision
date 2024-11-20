VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10320
   LinkTopic       =   "Form6"
   ScaleHeight     =   8925
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "GetImage"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   840
      Picture         =   "comlabel.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5280
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
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
      Left            =   480
      TabIndex        =   12
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   8640
      TabIndex        =   9
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   8640
      TabIndex        =   8
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   4560
      TabIndex        =   5
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "object finding"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5400
      MaskColor       =   &H00FFFFFF&
      Picture         =   "comlabel.frx":12FD
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "do it!!"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3720
      Picture         =   "comlabel.frx":1843
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Height          =   2655
      Left            =   4440
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   1
      Top             =   960
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   480
      Picture         =   "comlabel.frx":1EB3
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   0
      Top             =   960
      Width           =   3735
   End
   Begin VB.Shape Shape4 
      Height          =   2415
      Left            =   8520
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      Height          =   855
      Left            =   4440
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Shape Shape2 
      Height          =   855
      Left            =   240
      Top             =   4080
      Width           =   3975
   End
   Begin VB.Shape Shape1 
      Height          =   3135
      Left            =   240
      Top             =   720
      Width           =   8175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "maxk"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   11
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "max"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   10
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No of pixels"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Objects/Blobs"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   4320
      Width           =   3015
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim img(2000, 2000) As Long
Dim l As Long
Dim blob(100) As Long
Dim max As Long
Dim maxk As Long
Dim imax As Long, imin As Long, jmax As Long, jmin As Long
Dim min As Long
Dim grayscale As Byte
Dim val As Long
Dim gray As Long
Dim c As Double
Private Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Sub Command3_Click()
CommonDialog1.Filter = "Graphic Files (*.bmp;*.gif;*.jpg)| *.bmp;*.jpg"
CommonDialog1.InitDir = "D:\Astrovision\Images\test"
CommonDialog1.ShowOpen
getimage = CommonDialog1.FileName
Picture1.Picture = LoadPicture(getimage)
End Sub

Private Sub Command4_Click()
Unload Form6
Load Form1
Form1.Show
End Sub
Private Sub Form_Load()
c = 0
Form1.Visible = False
Form1.Cls
Form6.BackColor = RGB(70, 150, 240)
Picture1.Picture = LoadPicture("D:\Astrovision\Images\5.bmp")
Picture1.BackColor = vbBlack
Picture2.BackColor = vbBlack
End Sub

Private Sub Command1_Click()
' Get picture.................and find gray scale and do smoothing
c = 0
For j = 0 To Picture1.Picture.Width / 20
For i = 0 To Picture1.Picture.Height / 20
val = GetPixel(Picture1.hDC, i, j)
rr = val And 255
gg = (val \ 256) And 255
bb = (val \ 65536) And 255
grayscale = (rr / 3 + gg / 3 + bb / 3)
c = c + 1                                 ' no of pixels
If grayscale > 120 Then                   ' discard small points.!.!.!.!.!.!.!.!..
img(i, j) = grayscale
'img(i, j) = 255
Else
img(i, j) = 0
End If
Picture1.PSet (i, j), RGB(grayscale, grayscale, grayscale)
Next i
Next j
Text1.Text = c
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'smoothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''
For i = 1 To Picture1.Picture.Height / 10
For j = 1 To Picture1.Picture.Width / 10
If img(i, j) > 0 Then
img(i - 1, j) = 0.3 * (img(i, j) + img(i, j + 1) + img(i, j + 2) + img(i + 1, j) + img(i + 1, j + 1) + img(i + 1, j + 2) + img(i + 2, j) + img(i + 2, j + 1) + img(i + 2, j + 2)) / 9
img(i - 1, j - 1) = 0.5 * (img(i, j) + img(i, j + 1) + img(i, j + 2) + img(i + 1, j) + img(i + 1, j + 1) + img(i + 1, j + 2) + img(i + 2, j) + img(i + 2, j + 1) + img(i + 2, j + 2)) / 9
img(i, j - 1) = 0.3 * (img(i, j) + img(i, j + 1) + img(i, j + 2) + img(i + 1, j) + img(i + 1, j + 1) + img(i + 1, j + 2) + img(i + 2, j) + img(i + 2, j + 1) + img(i + 2, j + 2)) / 9
img(i + 1, j - 1) = 0.5 * (img(i, j) + img(i, j + 1) + img(i, j + 2) + img(i + 1, j) + img(i + 1, j + 1) + img(i + 1, j + 2) + img(i + 2, j) + img(i + 2, j + 1) + img(i + 2, j + 2)) / 9
img(i + 1, j) = 0.3 * (img(i, j) + img(i, j + 1) + img(i, j + 2) + img(i + 1, j) + img(i + 1, j + 1) + img(i + 1, j + 2) + img(i + 2, j) + img(i + 2, j + 1) + img(i + 2, j + 2)) / 9
img(i + 1, j + 1) = 0.5 * (img(i, j) + img(i, j + 1) + img(i, j + 2) + img(i + 1, j) + img(i + 1, j + 1) + img(i + 1, j + 2) + img(i + 2, j) + img(i + 2, j + 1) + img(i + 2, j + 2)) / 9
img(i, j + 1) = 0.3 * (img(i, j) + img(i, j + 1) + img(i, j + 2) + img(i + 1, j) + img(i + 1, j + 1) + img(i + 1, j + 2) + img(i + 2, j) + img(i + 2, j + 1) + img(i + 2, j + 2)) / 9
img(i - 1, j + 1) = 0.5 * (img(i, j) + img(i, j + 1) + img(i, j + 2) + img(i + 1, j) + img(i + 1, j + 1) + img(i + 1, j + 2) + img(i + 2, j) + img(i + 2, j + 1) + img(i + 2, j + 2)) / 9
img(i, j) = 6.4 * (img(i, j) + img(i, j + 1) + img(i, j + 2) + img(i + 1, j) + img(i + 1, j + 1) + img(i + 1, j + 2) + img(i + 2, j) + img(i + 2, j + 1) + img(i + 2, j + 2)) / 9
Else
img(i, j) = 0
End If
Picture1.PSet (i, j), RGB(img(i, j), img(i, j), img(i, j))
Next j
Next i
'''''''''''''''''''''''''''''''''''''''''''''''''''''''successful..............
End Sub
Private Sub Command2_Click()
Picture1.BackColor = vbBlack
Picture2.BackColor = vbBlack
Picture2.Picture = Picture1.Picture
' img (2000,2000) has the picture.......now label it
l = 0
For j = 1 To Picture1.Picture.Height / 20
For i = 1 To Picture1.Picture.Width / 20
If img(i, j) > 20 Then
        If img(i - 1, j) = 0 And img(i - 1, j - 1) = 0 And img(i, j - 1) = 0 And img(i + 1, j - 1) = 0 Then
            l = l + 1
            img(i, j) = l
            'l = l + 1
        Else
           'findmin:
            min = l
            img(i, j) = min
        End If
Else
img(i, j) = 0
End If
Next i
Next j

Text2.Text = l                   ' no of pixels in Pic 1

'pass 2 ok @@@@@@@@@@@@@@@@@@@@@@@@@@@

For j = 1 To Picture1.Picture.Height / 20
For i = 1 To Picture1.Picture.Width / 20
If img(i, j) > 0 And img(i, j) < 30 Then
    If img(i, j) < img(i - 1, j) Then
    img(i - 1, j) = img(i, j)
    End If
    If img(i, j) < img(i - 1, j - 1) Then
    img(i - 1, j - 1) = img(i, j)
    End If
    If img(i, j) < img(i, j - 1) Then
    img(i, j - 1) = img(i, j)
    End If
    If img(i, j) < img(i + 1, j - 1) Then
    img(i + 1, j - 1) = img(i, j)
    End If
    ' 4
    If img(i, j) < img(i + 1, j) Then
    img(i + 1, j) = img(i, j)
    End If
    If img(i, j) < img(i + 1, j + 1) Then
    img(i + 1, j + 1) = img(i, j)
    End If
    If img(i, j) < img(i, j + 1) Then
    img(i, j + 1) = img(i, j)
    End If
    If img(i, j) < img(i - 1, j + 1) Then
    img(i - 1, j + 1) = img(i, j)
    End If
End If
Next i
Next j

' boundary box ......for blobs with pixels >3000 + white
For k = 1 To 100
' wastage of looooop..................
blob(k) = 0
imin = 1000
jmin = 1000
imax = 0
jmax = 0
For i = 1 To Picture1.Picture.Width / 30
For j = 1 To Picture1.Picture.Height / 30
If img(i, j) = k Then
    blob(k) = blob(k) + 1
         If img(i, j) = k Then
            If i > imax Then
            imax = i
            End If
        End If
        If img(i, j) = k Then
            If j > jmax Then
            jmax = j
            End If
        End If
        If img(i, j) = k Then
            If j < jmin Then
            jmin = j
            End If
            End If
        If img(i, j) = k Then
            If i < imin Then
            imin = i
            End If
        End If
Else
End If
Next j
Next i
        Picture2.Line (imin, jmin)-(imax, jmin), vbWhite
        Picture2.Line (imax, jmin)-(imax, jmax), vbWhite
        Picture2.Line (imax, jmax)-(imin, jmax), vbWhite
        Picture2.Line (imin, jmax)-(imin, jmin), vbWhite
Next k
''''''''''''''''''''''''''''''''
max = blob(1)
maxk = 1
' boundary box ......for the largest blob with + RED
For k = 2 To 100
    If blob(k) > max Then
     maxk = k
     max = blob(maxk)
    End If
Next k
    Text3.Text = max
    Text4.Text = maxk
If max > 300 And max < 20000 Then      'Identify----Moon by size (size filter)
imin = 1000
jmin = 1000
imax = 0
jmax = 0
    For i = 1 To Picture1.Picture.Width / 20
    For j = 1 To Picture1.Picture.Height / 20
If img(i, j) = maxk Then
        If img(i, j) > 0 Then
            If i > imax Then
            imax = i
            End If
        End If
        If img(i, j) > 0 Then
            If j > jmax Then
            jmax = j
            End If
        End If
        If img(i, j) > 0 Then
            If j < jmin Then
            jmin = j
            End If
            End If
        If img(i, j) > 0 Then
            If i < imin Then
            imin = i
            End If
        End If
Else
'Picture2.PSet (i, j), vbBlack   ' for rest of picture PIC-BLOB
End If
    Next j
    Next i
Picture2.Line (imin, jmin)-(imax, jmin), vbRed
Picture2.Line (imax, jmin)-(imax, jmax), vbRed
Picture2.Line (imax, jmax)-(imin, jmax), vbRed
Picture2.Line (imin, jmax)-(imin, jmin), vbRed
End If
End Sub
Private Sub Form6_Terminate()
End
End Sub
