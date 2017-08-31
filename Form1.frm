VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MULTI CAL"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13815
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   13815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "&Clear All"
      BeginProperty Font 
         Name            =   "Harry P"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   24
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   405
      Left            =   8520
      TabIndex        =   20
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Calculate"
      BeginProperty Font 
         Name            =   "Harry P"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   19
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   10320
      TabIndex        =   16
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6600
      TabIndex        =   15
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   3000
      TabIndex        =   13
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&About"
      BeginProperty Font 
         Name            =   "Harry P"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      MaskColor       =   &H0000FF00&
      TabIndex        =   10
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Harry P"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   9
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Calculate"
      BeginProperty Font 
         Name            =   "Harry P"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "MS-1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   3600
      TabIndex        =   26
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label15 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   4680
      TabIndex        =   25
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "Simple Rule Movement"
      Height          =   255
      Left            =   1800
      TabIndex        =   23
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Momentum"
      Height          =   255
      Left            =   9240
      TabIndex        =   22
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Kg MS-1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   11160
      TabIndex        =   21
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000004&
      Caption         =   "MS-1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   12840
      TabIndex        =   18
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Kg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   9120
      TabIndex        =   17
      Top             =   480
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   6240
      X2              =   6240
      Y1              =   240
      Y2              =   4320
   End
   Begin VB.Label Label9 
      Caption         =   "KMH-1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Valocity"
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "MS-1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Speed"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Time"
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Distance / Distancement"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    num = Val(Text1.Text / Text2.Text)
    num = Text1.Text / Text2.Text
    Text3.Text = num
    
    Text4.Text = Text3.Text / Text2.Text
    
    Text5.Text = Text3.Text * 1000 / 3600
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
A = MsgBox("Made By A.M Shashika Theekshana Dayarathna Adhikari", vbOKOnly, "About")
End Sub


Private Sub Command4_Click()

     
     num = Val(Text6.Text * Text7.Text)
     num = Text6.Text * Text7.Text
     Text8.Text = num
End Sub


Private Sub Command5_Click()
'blank out the text boxs
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
End Sub
