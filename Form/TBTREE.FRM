VERSION 5.00
Begin VB.Form TBTree 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TracheoBroncialTree"
   ClientHeight    =   5820
   ClientLeft      =   90
   ClientTop       =   825
   ClientWidth     =   9135
   Icon            =   "TBTREE.frx":0000
   LinkTopic       =   "TracheoBroncialTree"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "TBTREE.frx":08CA
   ScaleHeight     =   388
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   609
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label rllb7 
      BackStyle       =   0  'Transparent
      Height          =   210
      Index           =   1
      Left            =   2730
      TabIndex        =   29
      Top             =   3840
      Width           =   405
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   405
      Left            =   2670
      TabIndex        =   28
      Top             =   4215
      Width           =   390
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   765
      Left            =   2760
      TabIndex        =   27
      Top             =   4650
      Width           =   840
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   3090
      TabIndex        =   26
      Top             =   3525
      Width           =   360
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   4395
      TabIndex        =   25
      Top             =   1845
      Width           =   810
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   2010
      TabIndex        =   24
      Top             =   1770
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   2565
      TabIndex        =   23
      Top             =   2460
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   420
      Left            =   2685
      TabIndex        =   22
      Top             =   870
      Width           =   1035
   End
   Begin VB.Label rll810 
      BackStyle       =   0  'Transparent
      Height          =   420
      Left            =   2955
      TabIndex        =   21
      Top             =   3825
      Width           =   540
   End
   Begin VB.Label rll 
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   3165
      TabIndex        =   20
      Top             =   3255
      Width           =   360
   End
   Begin VB.Label lms 
      BackStyle       =   0  'Transparent
      Height          =   825
      Left            =   5115
      TabIndex        =   19
      Top             =   2400
      Width           =   690
   End
   Begin VB.Label rms 
      BackStyle       =   0  'Transparent
      Height          =   600
      Left            =   3465
      TabIndex        =   18
      Top             =   2115
      Width           =   495
   End
   Begin VB.Label Proximal 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   3915
      TabIndex        =   17
      Top             =   -45
      Width           =   660
   End
   Begin VB.Label Distal 
      BackStyle       =   0  'Transparent
      Height          =   705
      Left            =   3990
      TabIndex        =   16
      Top             =   1740
      Width           =   660
   End
   Begin VB.Label lllb10 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   2910
      TabIndex        =   15
      Top             =   3360
      Width           =   285
   End
   Begin VB.Label llb9 
      BackStyle       =   0  'Transparent
      Height          =   270
      Left            =   3375
      TabIndex        =   14
      Top             =   3495
      Width           =   195
   End
   Begin VB.Label lllb8 
      BackStyle       =   0  'Transparent
      Height          =   1320
      Left            =   5280
      TabIndex        =   13
      Top             =   4275
      Width           =   2835
   End
   Begin VB.Label lllb6 
      BackStyle       =   0  'Transparent
      Height          =   405
      Left            =   5865
      TabIndex        =   12
      Top             =   3525
      Width           =   2340
   End
   Begin VB.Label lulb4 
      BackStyle       =   0  'Transparent
      Height          =   705
      Left            =   6210
      TabIndex        =   11
      Top             =   2730
      Width           =   1995
   End
   Begin VB.Label lulb12 
      BackStyle       =   0  'Transparent
      Height          =   1485
      Left            =   5445
      TabIndex        =   10
      Top             =   720
      Width           =   2970
   End
   Begin VB.Label rllb10 
      BackStyle       =   0  'Transparent
      Height          =   390
      Left            =   885
      TabIndex        =   9
      Top             =   5190
      Width           =   705
   End
   Begin VB.Label rllb9 
      BackStyle       =   0  'Transparent
      Height          =   645
      Left            =   825
      TabIndex        =   8
      Top             =   4395
      Width           =   1740
   End
   Begin VB.Label rllb8 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   855
      TabIndex        =   7
      Top             =   4050
      Width           =   1005
   End
   Begin VB.Label rllb7 
      BackStyle       =   0  'Transparent
      Height          =   345
      Index           =   0
      Left            =   825
      TabIndex        =   6
      Top             =   3675
      Width           =   780
   End
   Begin VB.Label rllb6 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   840
      TabIndex        =   5
      Top             =   3270
      Width           =   1005
   End
   Begin VB.Label rmlb5 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   750
      TabIndex        =   4
      Top             =   2730
      Width           =   1005
   End
   Begin VB.Label rmlb4 
      BackStyle       =   0  'Transparent
      Height          =   390
      Left            =   735
      TabIndex        =   3
      Top             =   2280
      Width           =   1035
   End
   Begin VB.Label RULB3 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   705
      TabIndex        =   2
      Top             =   1590
      Width           =   1365
   End
   Begin VB.Label RULB2 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   705
      TabIndex        =   1
      Top             =   1185
      Width           =   1005
   End
   Begin VB.Label rulb1 
      BackStyle       =   0  'Transparent
      Height          =   420
      Left            =   720
      TabIndex        =   0
      Top             =   705
      Width           =   1035
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "TBTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' All variables must be declared.
'Option Explicit

' Declare variables to hold the default
' width and height of the form.
Dim FileName As String

Private Sub Findings_Click()

End Sub


Private Sub Distal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\Distal.bmp")
Form1.Show
End Sub

Private Sub Form_Load()
Load Form1
End Sub

Private Sub mnuAbout_Click()
   
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Visible = False
Me.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
Unload Form1
Set TBTree = Nothing
Set Form1 = Nothing
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\rulb1.bmp")
Form1.Show

End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\rulb3.bmp")
Form1.Show

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\rulb2.bmp")
Form1.Show

End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\lmsluln.bmp")
Form1.Show

End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\rml1.bmp")
Form1.Show

End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\rllb10.bmp")
Form1.Show

End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\rllb9.bmp")
Form1.Show

End Sub

Private Sub llb9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\rmlb5.bmp")
Form1.Show

End Sub

Private Sub lllb10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\rmlb4.bmp")
Form1.Show

End Sub

Private Sub lllb6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\lllb6.bmp")
Form1.Show

End Sub

Private Sub lllb8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\Lllb8910.bmp")
Form1.Show

End Sub

Private Sub lms_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\lul.bmp")
Form1.Show


End Sub

Private Sub lulb12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\Lul1-23.bmp")
Form1.Show


End Sub

Private Sub lulb4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\lul45.bmp")
Form1.Show

End Sub

Private Sub mnuExit_Click()
   
   ' Terminate the program.
   Unload Me

End Sub

Private Sub mnuLarge_Click()

End Sub


Private Sub mnuMap_Click()

End Sub

Private Sub Proximal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\vocal.bmp")
Form1.Show

End Sub

Private Sub rll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\rll.bmp")
Form1.Show

End Sub

Private Sub rll810_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\rll8-10.bmp")
Form1.Show

End Sub

Private Sub rllb10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\rllb10.bmp")
Form1.Show

End Sub

Private Sub rllb6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\rllb6.bmp")
Form1.Show

End Sub

Private Sub rllb7_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\rllb8.bmp")
Form1.Show

End Sub

Private Sub rllb8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\rllb8.bmp")
Form1.Show

End Sub

Private Sub rllb9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\rllb9.bmp")
Form1.Show

End Sub

Private Sub rmlb4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\rmlb4.bmp")
Form1.Show

End Sub

Private Sub rmlb5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\rmlb5.bmp")
Form1.Show

End Sub

Private Sub rms_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\rul1.bmp")
Form1.Show


End Sub

Private Sub rulb1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\rulb1.bmp")
Form1.Show

End Sub

Private Sub RULB2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\rulb2.bmp")
Form1.Show

End Sub

Private Sub RULB3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Image1.Picture = LoadPicture(App.Path & "\Images\rulb3.bmp")
Form1.Show

End Sub
