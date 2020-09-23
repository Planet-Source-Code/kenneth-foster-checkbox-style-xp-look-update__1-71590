VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2820
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   2820
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Enabled"
      Height          =   330
      Left            =   1890
      TabIndex        =   9
      Top             =   2355
      Width           =   900
   End
   Begin Project1.ToggOpt ToggOpt6 
      Height          =   300
      Left            =   255
      TabIndex        =   7
      Top             =   1845
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   529
      Caption         =   "Test 6"
      Align           =   1
      Style           =   2
      OffSet          =   -5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
   End
   Begin Project1.ToggOpt ToggOpt5 
      Height          =   285
      Left            =   255
      TabIndex        =   6
      Top             =   1530
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   503
      Caption         =   "Test 5"
      Value           =   1
      Style           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.ToggOpt ToggOpt4 
      Height          =   330
      Left            =   255
      TabIndex        =   4
      Top             =   1125
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   582
      Caption         =   "Test 4"
      Align           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
   End
   Begin Project1.ToggOpt ToggOpt3 
      Height          =   300
      Left            =   255
      TabIndex        =   3
      Top             =   795
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   529
      Caption         =   "Test 3"
      Value           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.ToggOpt ToggOpt2 
      Height          =   315
      Left            =   255
      TabIndex        =   2
      Top             =   405
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "Test 2"
      ForeColor       =   12582912
      Align           =   1
      Style           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
   End
   Begin Project1.ToggOpt ToggOpt1 
      Height          =   285
      Left            =   255
      TabIndex        =   0
      Top             =   90
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      Caption         =   "Test 1"
      Style           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Checked"
      Height          =   225
      Left            =   1860
      TabIndex        =   8
      Top             =   1575
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "ON"
      Height          =   225
      Left            =   1875
      TabIndex        =   5
      Top             =   855
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "NO"
      Height          =   210
      Left            =   1875
      TabIndex        =   1
      Top             =   135
      Width           =   390
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   ToggOpt2.Enabled = Not ToggOpt2.Enabled
   ToggOpt4.Enabled = Not ToggOpt4.Enabled
   ToggOpt6.Enabled = Not ToggOpt6.Enabled
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub

Private Sub ToggOpt1_Click()
If ToggOpt1.Value = 0 Then
   Label1.Caption = "NO"
Else
   Label1.Caption = "YES"
End If
End Sub

Private Sub ToggOpt3_Click()
If ToggOpt3.Value = 0 Then
   Label2.Caption = "OFF"
Else
   Label2.Caption = "ON"
End If
End Sub

Private Sub ToggOpt5_Click()
If ToggOpt5.Value = 0 Then
   Label3.Caption = "Unchecked"
Else
   Label3.Caption = "Checked"
End If
End Sub
