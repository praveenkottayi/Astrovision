VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ASTRO VISION -mOOn Tracker series 1.0"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13275
   LinkTopic       =   "Form1"
   Picture         =   "Astrovision.frx":0000
   ScaleHeight     =   10155
   ScaleWidth      =   13275
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   7920
      Top             =   9480
   End
   Begin VB.CommandButton Command6 
      Height          =   2295
      Left            =   3600
      Picture         =   "Astrovision.frx":088F
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LIVE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   3600
      MaskColor       =   &H00404040&
      Picture         =   "Astrovision.frx":24AF
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Caption         =   "PTZ Features"
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
      Left            =   360
      Picture         =   "Astrovision.frx":37AC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7680
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Moon this Month !"
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
      Left            =   360
      Picture         =   "Astrovision.frx":53FB
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MOTION DETECTION MODE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      Picture         =   "Astrovision.frx":68FC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OBJECT SEARCH MODE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      Picture         =   "Astrovision.frx":85C8
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label10 
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
      Left            =   9000
      TabIndex        =   15
      Top             =   9480
      Width           =   375
   End
   Begin VB.Label Label9 
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
      Left            =   10200
      TabIndex        =   14
      Top             =   9480
      Width           =   375
   End
   Begin VB.Label Label8 
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
      Left            =   11400
      TabIndex        =   13
      Top             =   9480
      Width           =   1455
   End
   Begin VB.Shape Shape8 
      Height          =   615
      Left            =   8760
      Top             =   9360
      Width           =   4335
   End
   Begin VB.Label Label7 
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
      Left            =   9480
      TabIndex        =   12
      Top             =   9480
      Width           =   495
   End
   Begin VB.Label Label6 
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
      Left            =   10680
      TabIndex        =   11
      Top             =   9480
      Width           =   495
   End
   Begin VB.Shape Shape7 
      Height          =   375
      Left            =   11280
      Top             =   9480
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      Height          =   375
      Left            =   8880
      Top             =   9480
      Width           =   495
   End
   Begin VB.Shape Shape5 
      Height          =   375
      Left            =   10080
      Top             =   9480
      Width           =   495
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
      Left            =   9000
      TabIndex        =   10
      Top             =   9480
      Width           =   375
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
      Left            =   10200
      TabIndex        =   9
      Top             =   9480
      Width           =   375
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
      Left            =   11400
      TabIndex        =   8
      Top             =   9480
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   8760
      Top             =   9360
      Width           =   4335
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
      Left            =   9480
      TabIndex        =   7
      Top             =   9480
      Width           =   495
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
      Left            =   10680
      TabIndex        =   6
      Top             =   9480
      Width           =   495
   End
   Begin VB.Shape Shape2 
      Height          =   375
      Left            =   11280
      Top             =   9480
      Width           =   1695
   End
   Begin VB.Shape Shape3 
      Height          =   375
      Left            =   8880
      Top             =   9480
      Width           =   495
   End
   Begin VB.Shape Shape4 
      Height          =   375
      Left            =   10080
      Top             =   9480
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim theWebSite As String
'Option Explicit
Private Sub Command1_Click()
Unload Form1
Load Form2
Form2.Show
End Sub
Private Sub Command2_Click()
Unload Form1
Load Form6
Form6.Show
End Sub
Private Sub Command3_Click()
Load Form4
Form4.Show
Unload Form1
End Sub

Private Sub Command4_Click()

'theWebSite = "http://www.moonconnection.com/moon_phases_calendar.phtml"
'Call Shell("explorer.exe " & theWebSite, vbNormalFocus)

End Sub

Private Sub Command5_Click()
Unload Form1
Load Form3
Form3.Show
End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Form_Load()
Path = "D:\Astrovision\Images\8.jpg"
'Path = "C:\Documents and Settings\Administrator\Desktop\Astrovision\Images\8.jpg"
Form1.Picture = LoadPicture(Path)
Label10.Caption = Hour(Now)
Label9.Caption = Minute(Now)
Label8.Caption = DateValue(Now)
Timer1.Enabled = True
Timer1.Interval = 1000
End Sub
Private Sub Form1_Terminate()
End
End Sub
Private Sub Timer1_Timer()
Label10.Caption = Hour(Now)
Label9.Caption = Minute(Now)
End Sub
