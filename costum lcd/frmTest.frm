VERSION 5.00
Begin VB.Form frmTest 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin TestPrj.LCD LCD6 
      Height          =   225
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   397
   End
   Begin TestPrj.LCD LCD5 
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   397
   End
   Begin TestPrj.LCD LCD4 
      Height          =   225
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   397
   End
   Begin TestPrj.LCD LCD3 
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   397
   End
   Begin TestPrj.LCD LCD2 
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   397
   End
   Begin TestPrj.LCD LCD1 
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   397
   End
   Begin TestPrj.LCD LCD 
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   397
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   4440
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "you can create your own display style"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   3000
      Picture         =   "frmTest.frx":0000
      Top             =   240
      Width           =   1650
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
LCD1.LoadLCD App.Path & "\small1.bmp", 5, 9
LCD2.LoadLCD App.Path & "\small2.bmp", 6, 9
LCD3.LoadLCD App.Path & "\blue1.bmp", 10, 15
LCD4.LoadLCD App.Path & "\blue2.bmp", 20, 27
LCD5.LoadLCD App.Path & "\old1.bmp", 19, 27
LCD6.LoadLCD App.Path & "\old2.bmp", 19, 27
End Sub

Private Sub Timer1_Timer()
LCD.Caption = Format(Time, "hmm:ss")
LCD1.Caption = Format(Time, "hh:mm:ss")
LCD2.Caption = Format(Time, "hh:mm:ss")
LCD3.Caption = Format(Time, "hhmm:ss")
LCD4.Caption = Format(Time, "sshh:mm")
LCD5.Caption = Format(Time, "hh:ss")
LCD6.Caption = Format(Time, "mm:ss")
End Sub
