VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "BLOB labeling N bounding"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13485
   LinkTopic       =   "Form1"
   ScaleHeight     =   545
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   899
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   3975
      Left            =   2880
      TabIndex        =   10
      Top             =   4080
      Width           =   2535
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         Height          =   2655
         Left            =   120
         ScaleHeight     =   173
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   141
         TabIndex        =   20
         Top             =   600
         Width           =   2175
         Begin VB.PictureBox Picture6 
            Height          =   1455
            Left            =   240
            ScaleHeight     =   93
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   109
            TabIndex        =   23
            Top             =   240
            Width           =   1695
            Begin VB.PictureBox Picture4 
               AutoRedraw      =   -1  'True
               BorderStyle     =   0  'None
               Height          =   1575
               Left            =   0
               ScaleHeight     =   105
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   105
               TabIndex        =   24
               Top             =   0
               Width           =   1575
            End
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Stretch"
            Height          =   195
            Left            =   600
            TabIndex        =   22
            Top             =   2160
            Width           =   855
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1560
            ScaleHeight     =   23
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   23
            TabIndex        =   21
            Top             =   2040
            Visible         =   0   'False
            Width           =   375
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   19
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Object found :"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3855
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   5295
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   3550
         Left            =   120
         ScaleHeight     =   233
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   336
         TabIndex        =   5
         Top             =   240
         Width           =   5100
         Begin VB.VScrollBar VScroll1 
            Height          =   3255
            Left            =   4800
            TabIndex        =   7
            Top             =   0
            Width           =   255
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Top             =   3240
            Width           =   5055
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1815
            Left            =   0
            Picture         =   "Form1.frx":0000
            ScaleHeight     =   119
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   159
            TabIndex        =   8
            Top             =   0
            Width           =   2415
         End
      End
      Begin MSComDlg.CommonDialog cdl 
         Left            =   3000
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label3 
         Caption         =   "Input image (B/W only) :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   2655
      Begin VB.Frame Frame4 
         Height          =   2415
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   2415
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   16
            Text            =   "2"
            ToolTipText     =   "( Higher repetition more accuracy but slower - 4 usually enough )"
            Top             =   1200
            Width           =   495
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Debug mode "
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   15
            ToolTipText     =   "Very SLOW !!"
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Start labeling "
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   2175
         End
         Begin MSComctlLib.ProgressBar pbar 
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   2040
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "repetition:"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   17
            ToolTipText     =   "( Higher repetition more accuracy but slower - 4 usually enough )"
            Top             =   1200
            Width           =   1455
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   3360
         Width           =   2415
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Load Pict"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.TextBox Text1 
      Height          =   7695
      Left            =   5520
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   7815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim imageArray() As Byte
Dim imBW() As Byte
Dim oldX As Long, oldY As Long
Dim pNum As Long, pTot As Long

Private Sub Check2_Click()
If Check2.Value = 1 Then
        Call draw(True)
    Else
        Call draw(False)
    End If
End Sub

Private Sub Command1_Click()

Dim dib As New cDIB
Picture1.Cls
dib.GetImageData Picture1, imageArray
'-------------------------------------
Dim tmpheight As Long, tmpwidth As Long
tmpheight = Picture1.ScaleHeight
tmpwidth = Picture1.ScaleWidth
'-------------------------------------
Dim x As Long, y As Long
Dim xy() As Long
ReDim xy(0 To tmpwidth * 3)
For x = 0 To tmpwidth * 3
    xy(x) = x * 3
Next
'-------------------------------------
Dim temp As Long
ReDim imBW(tmpwidth - 1, tmpheight - 1)
'get image black and white
For x = 0 To tmpwidth - 1
For y = 0 To tmpheight - 1
    temp = imageArray(xy(x), y)
    If temp = 255 Then temp = 1
    imBW(x, y) = temp
Next
Next
'=======================[ MULTIPLE PASS PIXEL CONNECTIVITY LABELING ALGORITHM ]============================
'CHECK THE EXPLANATION OF ALGORITHM IN THE PDF FILE I INCLUDE IN ZIP FILE
'THERE IS OTHER FASTER WAY TO DO THIS PIXEL CONNECTIVTY LABELLING
'I JUZ IMPLEMENT ONE OF IT

Dim n As Long, min As Long, i As Long, j As Long, rep As Long
Dim label() As Long, mask(4) As Long, bscan As Long
ReDim label(tmpwidth - 1, tmpheight - 1)

If IsNumeric(Text2.Text) = False Then Exit Sub
rep = Text2.Text - 1

Dim amount As Long
amount = 1000
pbar.Max = Text2.Text + amount + 1
For j = 0 To rep
n = 1
fwdscan:
For x = 1 To tmpwidth - 2
For y = 1 To tmpheight - 2
    mask(0) = label(x - 1, y - 1)
    mask(1) = label(x, y - 1)
    mask(2) = label(x + 1, y - 1)
    mask(3) = label(x - 1, y)
    mask(4) = label(x, y)
    If imBW(x, y) = 1 Then
        temp = mask(0) Or mask(1) Or mask(2) Or mask(3)
        If temp = 0 Then
            label(x, y) = n: bscan = 1
            n = n + 1
        Else
            min = mask(0)
            For i = 1 To 4
                If min = 0 Then min = mask(i): GoTo cont
                If mask(i) < min And mask(i) <> 0 Then min = mask(i)
cont:
            Next
            label(x, y) = min: bscan = 1
        End If
    End If
Next
Next

backscan:
For x = 1 To tmpwidth - 2
For y = 1 To tmpheight - 2
    If label((tmpwidth - 1) - x, (tmpheight - 1) - y) <> 0 Then
    mask(0) = label((tmpwidth - 1) - (x - 1), (tmpheight - 1) - (y - 1))
    mask(1) = label((tmpwidth - 1) - x, (tmpheight - 1) - (y - 1))
    mask(2) = label((tmpwidth - 1) - (x + 1), (tmpheight - 1) - (y - 1))
    mask(3) = label((tmpwidth - 1) - (x - 1), (tmpheight - 1) - y)
    mask(4) = label((tmpwidth - 1) - x, (tmpheight - 1) - y)
    min = mask(0)
    For i = 1 To 4
        If min = 0 Then min = mask(i): GoTo cont2
        If mask(i) < min And mask(i) <> 0 Then min = mask(i)
cont2:
    Next
    
    label((tmpwidth - 1) - x, (tmpheight - 1) - y) = min
    End If
Next
Next

pbar.Value = j + 1
Next

finish:

Dim count() As Long
ReDim count(amount, 4)
Dim m As Long
For y = 0 To tmpheight - 1
For x = 0 To tmpwidth - 1
    count(label(x, y), 0) = count(label(x, y), 0) + 1 'set how many of label number 'no' has been got
    If count(label(x, y), 1) = 0 Then count(label(x, y), 1) = x: count(label(x, y), 2) = x: _
    count(label(x, y), 3) = y: count(label(x, y), 4) = y 'set all the min & max coordinate of x n y to min 1st..
    
    'update each coordinate
    If x < count(label(x, y), 1) And count(label(x, y), 1) <> 0 Then count(label(x, y), 1) = x 'update coordinate x min (if x < than the prev value)
    If x > count(label(x, y), 2) And count(label(x, y), 2) <> 0 Then count(label(x, y), 2) = x '
    If y < count(label(x, y), 3) And count(label(x, y), 3) <> 0 Then count(label(x, y), 3) = y '
    If y > count(label(x, y), 4) And count(label(x, y), 4) <> 0 Then count(label(x, y), 4) = y '
    
Next
Next

'======================[JUZ FOR DEBUG]==================
'PRINT OUT THE LABEL DONE TO THE TEXTBOX SO U CAN SEE IT

If Check1.Value = 1 Then
Dim str As String
str = ""
Text1.Text = ""
For y = 0 To tmpheight - 1
For x = 0 To tmpwidth - 1
    DoEvents
    If label(x, tmpheight - y - 1) < 10 Then
    str = str & "-" & "0" & label(x, tmpheight - y - 1)
    Else
    str = str & "-" & label(x, tmpheight - y - 1)
    End If
    If x = tmpwidth - 1 Then str = str & vbCrLf
Next
Next
Text1.Text = str
End If

'================[ STORE PICTURE OF BLOB FOUND ]============

Dim fname As String
n = 1
fname = App.Path & "\temp"

If Dir(App.Path & "\temp", vbDirectory) <> "" And Dir(App.Path & "\temp\*.bmp", vbNormal) <> "" Then
Kill App.Path & "\temp\*.bmp"
End If

If Dir(fname, vbDirectory) = "" Then MkDir (fname)
For i = 0 To amount
    If count(i, 0) <> 0 And count(i, 0) > 15 And i <> 0 Then
        Set Picture5.Picture = Picture3.Picture
        dib.CropImData Picture5, Picture1, imageArray, count(i, 1), count(i, 2), (tmpheight - 1) - count(i, 4), (tmpheight - 1) - count(i, 3), False
        fname = App.Path & "\temp\object" & n & ".bmp"
        SavePicture Picture5, fname
        n = n + 1
        
    End If
Next
Picture5.Picture = LoadPicture(App.Path & "\temp\object1.bmp")
If Check2.Value = 1 Then
    Call draw(True)
Else
    Call draw(False)
End If

pNum = 1 'picture counter
'===================[ BOUNDING BOX ]========================
'DRAW THE BOX AROUND DETECTED AND LABELED OBJECT



For i = 0 To amount
    If count(i, 0) <> 0 And count(i, 0) > 15 And i <> 0 Then 'if label found is not zero, > 5 and the label is not '0'
        m = m + 1
        
        Picture1.Line (count(i, 1), (tmpheight - 1) - count(i, 3))-(count(i, 2), (tmpheight - 1) - count(i, 4)), vbRed, B
       
    End If
     pbar.Value = pbar.Value + 1
Next

Label1.Caption = "object found :" & m

pTot = m

End Sub

Private Sub Command2_Click()
Dialog.Show , Me
Me.Enabled = False
End Sub

Private Sub Command3_Click()
cdl.CancelError = True
On Error GoTo err
cdl.Filter = "*.bmp|*.bmp"
cdl.InitDir = App.Path & "\images"
cdl.ShowOpen



If cdl.filename <> "" Then
Picture1 = LoadPicture(cdl.filename)
End If

'------------------[CHECK BW]----------------
Dim dib As New cDIB
 dib.GetImageData Picture1, imageArray
'-------------------------------------
Dim tmpheight As Long, tmpwidth As Long
tmpheight = Picture1.ScaleHeight
tmpwidth = Picture1.ScaleWidth
'-------------------------------------
Dim x As Long, y As Long
Dim xy() As Long
ReDim xy(0 To tmpwidth * 3)
For x = 0 To tmpwidth * 3
    xy(x) = x * 3
Next
'-------------------------------------
Dim temp As Long
ReDim imBW(tmpwidth - 1, tmpheight - 1)

For x = 0 To tmpwidth - 1
For y = 0 To tmpheight - 1
    temp = imageArray(xy(x), y)
    If temp <> 255 And temp <> 0 Then GoTo notbw
Next
Next

Command1.Enabled = True
err:
Exit Sub

notbw:
MsgBox "Not a bw image"
Command1.Enabled = False
Exit Sub
End Sub



Private Sub Command4_Click()
If pNum > 1 Then
    pNum = pNum - 1
    Picture5.Picture = LoadPicture(App.Path & "\temp\object" & pNum & ".bmp")
    If Check2.Value = 1 Then
        Call draw(True)
    Else
        Call draw(False)
    End If
End If
End Sub

Private Sub Command5_Click()
If pNum < pTot Then
    pNum = pNum + 1
    Picture5.Picture = LoadPicture(App.Path & "\temp\object" & pNum & ".bmp")
    If Check2.Value = 1 Then
        Call draw(True)
    Else
        Call draw(False)
    End If
End If
End Sub

Private Sub Form_Load()
Text1.Text = "this is for debug...show the label value"
Picture1_Resize
End Sub

Private Sub HScroll1_Change()
Picture1.Left = -HScroll1.Value * 50
End Sub

Private Sub Label4_Click()

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    If Button = 1 Then
        ' Debug.Print "test"
        Me.MousePointer = vbSizeAll
        'minus the scroller value with the diff of current mouse position and the refrnce mouse position
        If VScroll1.Enabled = True Then VScroll1.Value = VScroll1.Value - (y - oldY) / 50 'only update the scroller value
        If HScroll1.Enabled = True Then HScroll1.Value = HScroll1.Value - (x - oldX) / 50 'if scroller is enabled
        
        Else
        'if the left button is not down , update the reference _
        mouse position with current mouse position
        oldY = y
        oldX = x
        
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.MousePointer = vbArrow
End Sub

Private Sub Picture1_Resize()
HScroll1.Enabled = (Picture1.ScaleWidth > Picture2.ScaleWidth)
VScroll1.Enabled = (Picture1.ScaleHeight > Picture2.ScaleHeight)
VScroll1.Max = (Picture1.ScaleHeight - VScroll1.Height) / 50
HScroll1.Max = (Picture1.ScaleWidth - HScroll1.Width + VScroll1.Width) / 50
End Sub

Private Sub VScroll1_Change()
Picture1.Top = -VScroll1.Value * 50
End Sub

Public Function draw(stretch As Boolean)
Dim dib As New cDIB

    If stretch = True Then
        Picture4.Width = Picture6.Width
        Picture4.Height = Picture6.Height
    Else
        Picture4.Width = Picture5.ScaleWidth
        Picture4.Height = Picture5.ScaleHeight
    End If
    
   dib.CropImData Picture4, Picture5, imageArray, 0, Picture5.ScaleWidth, 0, Picture5.ScaleHeight, True
              

End Function
