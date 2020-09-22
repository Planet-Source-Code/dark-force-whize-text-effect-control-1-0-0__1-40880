VERSION 5.00
Object = "*\AWhize Text Effect Control.vbp"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Whize Text Effect Control 1.0.0"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin Project1.WhizeTextEffectControl WhizeTextEffectControl1 
      Height          =   660
      Left            =   75
      TabIndex        =   1
      Top             =   2070
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   1164
      BackColor       =   0
      Text            =   "Whize Text Effect Control 1.0.0"
      ForeColor       =   16777215
      EffectColor     =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Distance        =   100
      AutoSize        =   -1  'True
   End
   Begin Project1.WhizeTextEffectControl UserControl11 
      Height          =   660
      Left            =   90
      TabIndex        =   0
      Top             =   600
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   1164
      BackColor       =   0
      Text            =   "Whize Text Effect Control 1.0.0"
      Effect          =   1
      ForeColor       =   255
      EffectColor     =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Distance        =   100
      AutoSize        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Below is a text with the Shadow Effect applied"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   165
      TabIndex        =   3
      Top             =   1740
      Width           =   4710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Caption         =   "Below is a text with the 3D Effect applied"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   165
      TabIndex        =   2
      Top             =   270
      Width           =   4170
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
