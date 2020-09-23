VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Analog Meter Displays - By Max Seim - mlseim@mmm.com    12/08/00"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   4305
   ClientWidth     =   9675
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "analogmeter.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   1  'Arrow
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   645
   Begin VB.VScrollBar VScroll1 
      Height          =   975
      Index           =   5
      Left            =   9210
      Max             =   -90
      Min             =   90
      TabIndex        =   21
      Top             =   3600
      Value           =   1
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   975
      Index           =   4
      Left            =   9120
      Max             =   0
      Min             =   120
      TabIndex        =   20
      Top             =   1095
      Value           =   1
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2205
      Index           =   5
      Left            =   6810
      Picture         =   "analogmeter.frx":0442
      ScaleHeight     =   147
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   19
      Top             =   3045
      Width           =   2205
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2205
      Index           =   4
      Left            =   6795
      Picture         =   "analogmeter.frx":0BEB
      ScaleHeight     =   147
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   18
      Top             =   255
      Width           =   2205
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   975
      Index           =   3
      Left            =   6000
      Max             =   0
      Min             =   300
      TabIndex        =   11
      Top             =   3480
      Value           =   1
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   975
      Index           =   2
      Left            =   2760
      Max             =   0
      Min             =   100
      TabIndex        =   10
      Top             =   3480
      Value           =   1
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   975
      Index           =   1
      Left            =   6000
      Max             =   0
      Min             =   100
      TabIndex        =   9
      Top             =   1080
      Value           =   1
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   975
      Index           =   0
      Left            =   2640
      Max             =   0
      Min             =   360
      TabIndex        =   8
      Top             =   1080
      Value           =   1
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1185
      Index           =   3
      Left            =   3720
      Picture         =   "analogmeter.frx":145C
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   3
      Top             =   3240
      Width           =   2265
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1125
      Index           =   2
      Left            =   480
      Picture         =   "analogmeter.frx":1A9F
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   2
      Top             =   3240
      Width           =   2205
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1125
      Index           =   1
      Left            =   3720
      Picture         =   "analogmeter.frx":206C
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   1
      Top             =   840
      Width           =   2205
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2205
      Index           =   0
      Left            =   360
      Picture         =   "analogmeter.frx":2689
      ScaleHeight     =   147
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   0
      Top             =   240
      Width           =   2205
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Level Indication -  degs."
      Height          =   285
      Left            =   6615
      TabIndex        =   25
      Top             =   5370
      Width           =   2685
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Air Pressure"
      Height          =   255
      Left            =   7200
      TabIndex        =   24
      Top             =   2520
      Width           =   1785
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   5
      Left            =   9015
      TabIndex        =   23
      Top             =   4695
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   4
      Left            =   9000
      TabIndex        =   22
      Top             =   2235
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "0-300 Deg F"
      Height          =   375
      Left            =   4200
      TabIndex        =   17
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0-100 %"
      Height          =   375
      Left            =   1200
      TabIndex        =   16
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   3
      Left            =   6000
      TabIndex        =   15
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   14
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   1
      Left            =   6000
      TabIndex        =   13
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   12
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Temperature Level"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Valve Position"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Motor RPM"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Direction"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'                   Analog Meter Movements
'
'                   By:  Max Seim   mlseim@mmm.com
'
'                   Allows you to create your own analog meters using MSpaint.
'                   These images become the meter face.  Use the (analogmeter)
'                   subroutine to draw the hand (needle) automatically scaling
'                   the needle to the size of the Picture Box and position the
'                   needle to the engineering unit (value) you wish to display.
'
'                   Movement is smooth.  You can also vary the needle width, color,
'                   and needle length.  All of the meter attributes are sent to the
'                   subroutine with each update, allowing the needle color, (for example)
'                   to change as a value increases.
'
'                   Using the Picture1.Picture property, you can even load in a
'                   different meter face during runtime ... the meter needle and the
'                   meter face to not conflict with each other.
'
Dim sl As Integer
Dim r1 As Double
Dim r2 As Double
Const vbPI = 3.141592654
Const Deg2Rad = vbPI / 180 'Degrees to Radians
Dim cx As Integer, cy As Integer, r As Single
Dim s As Double

Private Sub Form_Load()
Form1.Left = (Screen.Width / 2) - (Form1.Width / 2)
Form1.Top = (Screen.Height / 2) - (Form1.Height / 2)
 'zero-out all the meters
 For X = 0 To 5
 VScroll1(X).Value = 0
 Next X
End Sub

Private Sub analogmeter(Index As Integer, Mtype As Integer, Emin As Double, Emax As Double, _
Mmin As Double, Mmax As Double, Handw As Integer, Color As String, Handl As Double, _
Value As Double)
'
' The meter movement is based off of a clock (0-60 minutes)
'
' The meter is assigned to a picture array (using Picture1).
' The Picture1 attributes are important.  Set Scalemode = Pixel (3)
'                                             Autosize = true
'                                             Autoredraw = true
'                                             Font Transparent = true
'                                             Border Style = your preference
'                                             Backcolor = your preference
'
' The Meter Face Image:
' I used MSpaint (the version that allows transparent .gif type format).
' The image is drawn so that the meter needle is centered, or centered at the
' bottom of the image.  Saved as a transparent .gif so that the background will
' take-on whatever Picture1.Backcolor you choose (see attribute above).
' When you load in the image into Picture1, the Picture1 Box will autosize to
' the .gif image (see attribute above).
'
' Meter Type 1 (center dial), places the zero at the bottom, 6 o'clock position.
' Meter Type 2 (half dial), places the zero at the left, 9 o'clock position.
' Meter Type 3 (center dial), Needle with extended arm, ball and arrow. (compass)
' Meter Type 4 (center dial), Full needle (level indicator type).
'
' Because not all meters start the needle at the bottom, you need to specify the
' Meter Minimum (MMin) and Maximum (MMax) to tell it where the zero and range
' offsets are at.  Example is the Air Pressure Meter.  Remember that the values
' are equivalent to the minutes of a clock.
'
' The Engineering units are the zero and range of the value you wish to display.
' If a meter has no numbers on it, you can simply pick 0 to 100.
'
' Index = Which Meter to Adjust
' Mtype = Which Type of Meter it is (1=center dial, 2=half dial, 3=pointer needle)
' EMin  = Engineering Units Zero (Minimum)
' EMax  = Engineering Units Span (Range)
' MMin  = Minimum point on the meter face (0-60)
' MMax  = Maximum point on the meter face (0-60)
' Handw = Thickness of dial hand (1,2,3)
' Color = Color of dial hand (vbRed, vbBlack)
' Handl = Length of the dial hand (0.1 is a good length)
' Value = The value to set the dial

' Determine whether dial hand is in center or edge.
' Also scale the dial hand to picture dimensions
If Mtype = 1 Then
Mdeg = 0 ' degrees (0-360)
cx = Picture1(Index).Width / 2
cy = Picture1(Index).Height / 2
End If
If Mtype = 2 Then
Mdeg = 270 ' degrees (0-360)
cx = Picture1(Index).Width / 2
cy = (Picture1(Index).Height - 2)
End If
If Mtype = 3 Then
Mdeg = 0 ' degrees (0-360)
cx = Picture1(Index).Width / 2
cy = Picture1(Index).Height / 2
End If
If Mtype = 4 Then
Mdeg = 0 ' degrees (0-360)
cx = Picture1(Index).Width / 2
cy = Picture1(Index).Height / 2
End If
' Scale the dial hand length
r = IIf(cx > cy, cy, cx) - 5
sl = r * Handl 'length of meter hand

' Scale the Engineering Units
r1 = Emax - Emin
r2 = Mmax - Mmin
s = ((r2 / r1) * Value) + Mmin
  
' Draw the dial hand
  sin_ = Sin((Mdeg - s * 6) * Deg2Rad) * (r - sl) + cx
  cos_ = Cos((Mdeg - s * 6) * Deg2Rad) * (r - sl) + cy
  '
  ' Special calculations for a Type 3 needle:
  '----------------------------------------------------------
  ' Setting the length of the "back of the needle"
  ' Default = half the length of the "front of the needle"
  ' To alter, change ((r - sl) / 2)  to a different ratio.
  If Mtype = 3 Then
  sin2_ = Sin((180 - s * 6) * Deg2Rad) * ((r - sl) / 1.5) + cx
  cos2_ = Cos((180 - s * 6) * Deg2Rad) * ((r - sl) / 1.5) + cy
  ar1 = Sin((45 - s * 6) * Deg2Rad) * ((r - sl) / 2) + cx
  ar2 = Cos((45 - s * 6) * Deg2Rad) * ((r - sl) / 2) + cy
  ar3 = Sin((315 - s * 6) * Deg2Rad) * ((r - sl) / 2) + cx
  ar4 = Cos((315 - s * 6) * Deg2Rad) * ((r - sl) / 2) + cy
  End If
  If Mtype = 4 Then
  sin2_ = Sin((180 - s * 6) * Deg2Rad) * (r - sl) + cx
  cos2_ = Cos((180 - s * 6) * Deg2Rad) * (r - sl) + cy
  End If
  oldcolor = vbBlack
  Picture1(Index).ForeColor = Color
  Picture1(Index).DrawWidth = Handw
  Picture1(Index).Cls
  Picture1(Index).Line (cx, cy)-(sin_, cos_)
  If Mtype = 3 Then ' Draw special features on Type 3 Needle:
  Picture1(Index).Line (cx, cy)-(sin2_, cos2_)
  Picture1(Index).Line (sin_, cos_)-(ar1, ar2) ' Set needle arrow head
  Picture1(Index).Line (sin_, cos_)-(ar3, ar4) ' Set needle arrow head
  Picture1(Index).FillColor = Color ' set FillColor
  Picture1(Index).FillStyle = 0 ' set FillStyle to SOLID
  Picture1(Index).Circle (sin2_, cos2_), 5 ' 5 = The size of the circle on the "back of needle".
  End If
  If Mtype = 4 Then ' Draw special features on Type 4 Needle:
  Picture1(Index).Line (cx, cy)-(sin2_, cos2_)
  End If
  Picture1(Index).ForeColor = oldcolor

End Sub

Private Sub VScroll1_Change(Index As Integer)
' Update Meter values
Label5(Index) = VScroll1(Index).Value
If Index = 0 Then ' Compass - Needle with arrow and ball end
analogmeter Index, 3, 0, 360, 0, 60, 4, vbRed, 0.2, VScroll1(Index).Value
End If
If Index = 1 Then ' Motor RPM Meter
analogmeter Index, 2, 0, 100, 0, 30, 3, vbBlue, 0.1, VScroll1(Index).Value
End If
If Index = 2 Then ' Valve Position Meter
analogmeter Index, 2, 0, 100, 0, 30, 2, vbBlack, 0.1, VScroll1(Index).Value
End If
If Index = 3 Then ' Temperature Meter
analogmeter Index, 2, 0, 300, 0, 30, 4, vbRed, 0.1, VScroll1(Index).Value
End If
If Index = 4 Then ' Air Pressure Gauge
analogmeter Index, 1, 0, 120, 7, 53, 1, vbBlack, 0.1, VScroll1(Index).Value
End If
If Index = 5 Then ' Level Indication
analogmeter Index, 4, -90, 90, 15, 45, 4, vbGreen, 0#, VScroll1(Index).Value
End If
End Sub

