VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "CheckBoxes"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   3480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame1"
      Height          =   735
      Left            =   720
      TabIndex        =   12
      Top             =   2760
      Width           =   1935
      Begin Project1.CheckBox CheckBox13 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         BackColor       =   16761024
         BackStyle       =   1
         BoxBackColor    =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Project1.CheckBox CheckBox4 
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BackColor       =   16777215
      BackStyle       =   1
      BoxBackColor    =   16761024
      BoxBorderDark   =   0
      BoxStyle        =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   1
   End
   Begin Project1.CheckBox CheckBox3 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BoxBackColor    =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.CheckBox CheckBox2 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BoxBackColor    =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   1
   End
   Begin Project1.CheckBox CheckBox1 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BoxBackColor    =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.CheckBox CheckBox5 
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BoxBackColor    =   12648384
      BoxBorderDark   =   0
      BoxStyle        =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.CheckBox CheckBox6 
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BoxBackColor    =   12632319
      BoxBorderDark   =   0
      BoxStyle        =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.CheckBox CheckBox7 
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BoxStyle        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MarkColor       =   16711680
      Value           =   1
   End
   Begin Project1.CheckBox CheckBox8 
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BoxStyle        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MarkColor       =   16711680
   End
   Begin Project1.CheckBox CheckBox9 
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BoxStyle        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MarkColor       =   16711680
   End
   Begin Project1.CheckBox CheckBox10 
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BoxStyle        =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MarkColor       =   255
   End
   Begin Project1.CheckBox CheckBox11 
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BoxStyle        =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MarkColor       =   255
      Value           =   1
   End
   Begin Project1.CheckBox CheckBox12 
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   2280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BoxStyle        =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MarkColor       =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


