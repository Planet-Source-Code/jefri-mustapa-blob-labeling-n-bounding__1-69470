VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   0  'None
   Caption         =   "Dialog Caption"
   ClientHeight    =   1950
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   4935
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4230
      Top             =   1320
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1035
      Left            =   240
      Picture         =   "Dialog.frx":0000
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   2
      Top             =   240
      Width           =   1500
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2010
      Left            =   120
      Picture         =   "Dialog.frx":4962
      ScaleHeight     =   130
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   285
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CommandButton OKButton 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   1500
   End
   Begin VB.PictureBox Picture3 
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1635
      ScaleWidth      =   4635
      TabIndex        =   3
      Top             =   120
      Width           =   4695
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   120
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   1680
         TabIndex        =   4
         Top             =   480
         Width           =   2880
      End
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4920
      X2              =   4920
      Y1              =   0
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   6360
      Y1              =   1940
      Y2              =   1940
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function BitBlt Lib "gdi32" _
    (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal dwRop As Long) As Long


Private Sub Form_Load()
Dim str As String
str = "Code by Jefri b Mustapa [sukiminna@yahoo.com]" & vbCr & vbCr _
      & "Please comment it if u find bugs or maybe can make it much better"
Label1.Caption = str
Label2.Caption = "IMAGE PROGRAMMING 4 FUN"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Enabled = True
'Gray.SetFocus

End Sub

Private Sub OKButton_Click()
Unload Me
End Sub



Private Sub Timer1_Timer()
Dim x As Long
DoEvents
Randomize
Picture2.Cls
x = Rnd * 20
'Debug.Print x
BitBlt Picture2.hDC, 7, 5, 8 * 10, 60, Picture1.hDC, x * 8, Picture1.ScaleHeight \ 2, vbSrcAnd
Picture2.Refresh

End Sub
