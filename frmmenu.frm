VERSION 5.00
Begin VB.Form frmmenu 
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1755
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   1755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Height          =   135
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   135
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   360
      X2              =   1320
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  'Inside Solid
      Index           =   0
      X1              =   360
      X2              =   1320
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   0
      Picture         =   "frmmenu.frx":0000
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Menu 1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Menu 1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Menu 1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Menu 1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Menu 1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   0
      Picture         =   "frmmenu.frx":0296
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1740
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bold
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
mnuitem Me, 0, "Open"
mnuitem Me, 1, "Close"
mnuitem Me, 2, "Print"
mnuitem Me, 3, "Add"
mnuitem Me, 4, "Exit"
'Make the form translucent...
Call MakeTranslucent(Me, &HFF00&)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1(bold).FontBold = False
End Sub

Private Sub Form_Resize()
If iRecursion Then Exit Sub

'Make the form translucent...
Call MakeTranslucent(Me, &HFF00&)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    'Move form
    Call DragForm(Me)
    'Make the form translucent...
    Call MakeTranslucent(Me, &HFF00&)
End If
End Sub

Private Sub Image2_Click()
End
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
Case 0
MsgBox Label1(Index).Caption
Case 1
MsgBox Label1(Index).Caption
Case 2
MsgBox Label1(Index).Caption
Case 3
MsgBox Label1(Index).Caption
Case 4
MsgBox Label1(Index).Caption
End Select
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1(Index).FontBold = True
bold = Index
End Sub
