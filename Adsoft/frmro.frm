VERSION 5.00
Begin VB.Form frmro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RELEASE ORDER"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   10395
   Begin VB.ComboBox cmbrono 
      Height          =   315
      Left            =   2063
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox txtclient 
      Height          =   375
      Left            =   2063
      MaxLength       =   50
      TabIndex        =   2
      Top             =   660
      Width           =   3855
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5555
      TabIndex        =   38
      ToolTipText     =   "Cancel Current RO"
      Top             =   6840
      Width           =   1215
   End
   Begin VB.TextBox txtlanguage 
      Height          =   315
      Left            =   6240
      MaxLength       =   15
      TabIndex        =   35
      Top             =   6240
      Width           =   2295
   End
   Begin VB.TextBox txtmaterial 
      Height          =   315
      Left            =   2520
      MaxLength       =   15
      TabIndex        =   34
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&New"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1695
      TabIndex        =   36
      ToolTipText     =   "Make and save new RO"
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7485
      TabIndex        =   39
      ToolTipText     =   "Close This Screen"
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3625
      TabIndex        =   37
      ToolTipText     =   "Print this RO"
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Index           =   2
      Left            =   360
      TabIndex        =   62
      Top             =   4680
      Width           =   9735
      Begin VB.TextBox txtppercent 
         Height          =   315
         Index           =   2
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   30
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtsize1 
         Height          =   315
         Index           =   2
         Left            =   7320
         MaxLength       =   5
         TabIndex        =   32
         Top             =   960
         Width           =   1065
      End
      Begin VB.TextBox txtsize2 
         Height          =   315
         Index           =   2
         Left            =   8535
         MaxLength       =   5
         TabIndex        =   33
         Top             =   960
         Width           =   1065
      End
      Begin VB.TextBox txtcaption 
         Height          =   315
         Index           =   2
         Left            =   120
         MaxLength       =   35
         TabIndex        =   28
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox cmbpublication 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox cmbedition 
         Height          =   315
         Index           =   2
         Left            =   2520
         TabIndex        =   25
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox cmbposition 
         Height          =   315
         Index           =   2
         Left            =   4920
         TabIndex        =   26
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox cmbpremium 
         Height          =   315
         Index           =   2
         Left            =   7320
         TabIndex        =   27
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtrate 
         Height          =   315
         Index           =   2
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   29
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtrdate 
         Height          =   315
         Index           =   2
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   31
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Pr. Percent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   3720
         TabIndex        =   87
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Col."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   8535
         TabIndex        =   82
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Cm."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   7335
         TabIndex        =   81
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   8400
         TabIndex        =   80
         Top             =   1020
         Width           =   120
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   69
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Publication"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   68
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Edition"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2520
         TabIndex        =   67
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   4920
         TabIndex        =   66
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   65
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Release Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   4920
         TabIndex        =   64
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Premium"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   7320
         TabIndex        =   63
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Index           =   1
      Left            =   360
      TabIndex        =   54
      Top             =   3240
      Width           =   9735
      Begin VB.TextBox txtppercent 
         Height          =   315
         Index           =   1
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   20
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtsize1 
         Height          =   315
         Index           =   1
         Left            =   7320
         MaxLength       =   5
         TabIndex        =   22
         Top             =   960
         Width           =   1065
      End
      Begin VB.TextBox txtsize2 
         Height          =   315
         Index           =   1
         Left            =   8535
         MaxLength       =   5
         TabIndex        =   23
         Top             =   960
         Width           =   1065
      End
      Begin VB.TextBox txtrdate 
         Height          =   315
         Index           =   1
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   21
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtrate 
         Height          =   315
         Index           =   1
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox cmbpremium 
         Height          =   315
         Index           =   1
         Left            =   7320
         TabIndex        =   17
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox cmbposition 
         Height          =   315
         Index           =   1
         Left            =   4920
         TabIndex        =   16
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox cmbedition 
         Height          =   315
         Index           =   1
         Left            =   2520
         TabIndex        =   15
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox cmbpublication 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtcaption 
         Height          =   315
         Index           =   1
         Left            =   120
         MaxLength       =   35
         TabIndex        =   18
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Pr. Percent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   3720
         TabIndex        =   86
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Col."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   8535
         TabIndex        =   79
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Cm."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   7335
         TabIndex        =   78
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   8400
         TabIndex        =   77
         Top             =   1020
         Width           =   120
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Premium"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   7320
         TabIndex        =   61
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Release Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4920
         TabIndex        =   60
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   59
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4920
         TabIndex        =   58
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Edition"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   57
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Publication"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   56
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   55
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Index           =   0
      Left            =   360
      TabIndex        =   45
      Top             =   1800
      Width           =   9735
      Begin VB.TextBox txtppercent 
         Height          =   315
         Index           =   0
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtsize2 
         Height          =   315
         Index           =   0
         Left            =   8520
         MaxLength       =   5
         TabIndex        =   13
         Top             =   960
         Width           =   1065
      End
      Begin VB.TextBox txtcaption 
         Height          =   315
         Index           =   0
         Left            =   120
         MaxLength       =   35
         TabIndex        =   8
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox cmbpublication 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox cmbedition 
         Height          =   315
         Index           =   0
         Left            =   2520
         TabIndex        =   5
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox cmbposition 
         Height          =   315
         Index           =   0
         Left            =   4920
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox cmbpremium 
         Height          =   315
         Index           =   0
         Left            =   7320
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtrate 
         Height          =   315
         Index           =   0
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtrdate 
         Height          =   315
         Index           =   0
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   11
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtsize1 
         Height          =   315
         Index           =   0
         Left            =   7305
         MaxLength       =   5
         TabIndex        =   12
         Top             =   960
         Width           =   1065
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Pr. Percent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3720
         TabIndex        =   85
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Col."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   8520
         TabIndex        =   76
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   8385
         TabIndex        =   75
         Top             =   1020
         Width           =   120
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   53
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Publication"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   52
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Edition"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   51
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   4920
         TabIndex        =   50
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   49
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Release Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   4920
         TabIndex        =   48
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Cm."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   7320
         TabIndex        =   47
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Premium"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   7320
         TabIndex        =   46
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.TextBox txtadmanager 
      Height          =   405
      Left            =   2063
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox txtdate 
      Height          =   375
      Left            =   7463
      MaxLength       =   10
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Release Order"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   3
      Left            =   6720
      TabIndex        =   84
      Top             =   1200
      Width           =   3165
   End
   Begin VB.Label lblcancel 
      Caption         =   "C A N C E L L E D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4320
      TabIndex        =   83
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lbldate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Ex. 26-03-1978)"
      Height          =   195
      Left            =   7500
      TabIndex        =   74
      Top             =   540
      Width           =   1170
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Language"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   73
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Material"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   72
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   71
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   70
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   44
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6270
      TabIndex        =   43
      Top             =   120
      Width           =   675
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Advertising Manager"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   510
      TabIndex        =   42
      Top             =   1080
      Width           =   1260
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Client"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   510
      TabIndex        =   41
      Top             =   660
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NO."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   510
      TabIndex        =   40
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3BDF9E9D0048"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Form"
Dim arrpubid()       As Long  'Publication id
Dim arreditionid0()  As Long  'Edition id
Dim arreditionid1()  As Long  'Edition id
Dim arreditionid2()  As Long  'Edition id
Dim arrpositionid0() As Long  'positions id
Dim arrpositionid1() As Long  'Position id
Dim arrpositionid2() As Long  'Position id
Dim arrpremiumid0()  As Long  'Premium id
Dim arrpremiumid1()  As Long  'Premium id
Dim arrpremiumid2()  As Long  'Premium id
'*********************************anish*******************************
Dim arrRoId() As Integer

Private Sub cmbedition_Click(Index As Integer)
If cmbedition(Index) = "" Then Exit Sub

Select Case Index
Case Is = 0
    addPosition arrpubid(cmbpublication(Index).ListIndex), arreditionid0(cmbedition(Index).ListIndex), CByte(Index)
Case Is = 1
    addPosition arrpubid(cmbpublication(Index).ListIndex), arreditionid1(cmbedition(Index).ListIndex), CByte(Index)
Case Is = 2
    addPosition arrpubid(cmbpublication(Index).ListIndex), arreditionid2(cmbedition(Index).ListIndex), CByte(Index)
End Select

End Sub

Private Sub cmbposition_Click(Index As Integer)
    Dim rec As New ADODB.Recordset
    If cmbposition(Index) = "" Then Exit Sub
    
    Select Case Index
    Case Is = 0
    rec.Open "select rate from ratecard where publication=" & arrpubid(cmbpublication(Index).ListIndex) & " and edition=" & arreditionid0(cmbedition(Index).ListIndex) & " and positions=" & arrpositionid0(cmbposition(Index).ListIndex), cn, adOpenStatic, adLockReadOnly
    Case Is = 1
    rec.Open "select rate from ratecard where publication=" & arrpubid(cmbpublication(Index).ListIndex) & " and edition=" & arreditionid1(cmbedition(Index).ListIndex) & " and positions=" & arrpositionid1(cmbposition(Index).ListIndex), cn, adOpenStatic, adLockReadOnly
    Case Is = 2
    rec.Open "select rate from ratecard where publication=" & arrpubid(cmbpublication(Index).ListIndex) & " and edition=" & arreditionid2(cmbedition(Index).ListIndex) & " and positions=" & arrpositionid2(cmbposition(Index).ListIndex), cn, adOpenStatic, adLockReadOnly
    End Select
    
'    adrate(Index) = rec!Rate
    txtrate(Index) = rec!Rate
    cmbpremium(Index) = "None"
End Sub

Private Sub cmbpremium_Click(Index As Integer)
    Dim rec As New ADODB.Recordset
    
    If cmbpremium(Index) = "" Then Exit Sub
    
    If cmbpremium(Index) = "None" Then
        txtppercent(Index) = ""
        Exit Sub
    End If
    
    Select Case Index
    Case Is = 0
    rec.Open "select premium_percent from premium where premium_id=" & arrpremiumid0(cmbpremium(Index).ListIndex - 1), cn, adOpenStatic, adLockReadOnly
    Case Is = 1
    rec.Open "select premium_percent from premium where premium_id=" & arrpremiumid1(cmbpremium(Index).ListIndex - 1), cn, adOpenStatic, adLockReadOnly
    Case Is = 2
    rec.Open "select premium_percent from premium where premium_id=" & arrpremiumid2(cmbpremium(Index).ListIndex - 1), cn, adOpenStatic, adLockReadOnly
    End Select
    
    txtppercent(Index) = rec!premium_percent
'    txtrate(Index) = adrate(Index) + adrate(Index) * (rec!premium_percent / 100)
 
 End Sub

Private Sub cmbpublication_Click(Index As Integer)
If cmbpublication(Index) = "" Then Exit Sub
addEdition arrpubid(cmbpublication(Index).ListIndex), CByte(Index)
addPremium arrpubid(cmbpublication(Index).ListIndex), CByte(Index)
End Sub

Private Sub cmbrono_Click()
Dim rec As New ADODB.Recordset

ClearControls
If cmbrono = "" Then Exit Sub
rec.Open "select * from romaster where rono=" & Split(cmbrono, "/")(0) & " and finyear=" & Val("20" & Split(cmbrono, "/")(1)), cn, adOpenStatic, adLockReadOnly
txtdate = Format(rec("rodate"), "dd-mm-yyyy")
txtclient = rec("client")
txtadmanager = rec("admanager")
txtmaterial = rec("material")
txtlanguage = rec("language")
lblcancel.Visible = rec("cancelled")
cmdcancel.Enabled = Not (rec("cancelled"))
cmdprint.Enabled = Not (rec("cancelled"))

'**************************Anish**********************************************************

Dim recDetail As New ADODB.Recordset
'getting rodetail fields........
recDetail.Open "select * from  rodetail, romaster where romaster.romaster_ID=rodetail.romaster_Id and  romaster.romaster_ID=" & rec("romaster_ID") & "", cn, adOpenStatic, adLockReadOnly

'displaying records.......
i = 0
While recDetail.EOF = False
    cmbpublication(i) = recDetail("publication")
    cmbedition(i) = recDetail("edition")
    cmbposition(i) = recDetail("position")
    txtcaption(i) = recDetail("caption")
    cmbpremium(i) = recDetail("premium_name")
    txtrate(i) = recDetail("rate")
    txtrdate(i) = Format(recDetail("releasedate"), "dd-mm-yyyy")
    txtsize1(i) = recDetail("size1")
    If recDetail("size2") > 0 Then txtsize2(i) = recDetail("size2")
    If recDetail("premium_percent") <> 0 Then txtppercent(i) = recDetail("premium_percent")
    i = i + 1
    recDetail.MoveNext
Wend
'********************************************************************************

'''''rec.Close
'''''rec.Open "select * from rodetail where rono=" & Split(cmbrono, "/")(0) & " and finyear=" & Val("20" & Split(cmbrono, "/")(1)), cn, adOpenStatic, adLockReadOnly
'''''i = 0
'''''While rec.EOF = False
'''''cmbpublication(i) = rec("publication")
'''''cmbedition(i) = rec("edition")
'''''cmbposition(i) = rec("position")
'''''txtcaption(i) = rec("caption")
'''''cmbpremium(i) = rec("premium_name")
'''''txtrate(i) = rec("rate")
'''''txtrdate(i) = Format(rec("releasedate"), "dd-mm-yyyy")
'''''txtsize1(i) = rec("size1")
'''''If rec("size2") > 0 Then txtsize2(i) = rec("size2")
'''''If rec("premium_percent") <> 0 Then txtppercent(i) = rec("premium_percent")
'''''i = i + 1
'''''rec.MoveNext
'''''Wend
End Sub

Private Sub cmdcancel_Click()
Dim rec As New ADODB.Recordset, ask As Integer
On Error GoTo errorHandler
Dim str As String
'anish
rec.Open "select billno, finyear from billmaster where ref='r' and romaster_id=" & arrRoId(cmbrono.ListIndex) & "", cn, adOpenStatic, adLockReadOnly
'rec.Open "select billno, finyear from billmaster where ref='r' and refno='" & cmbrono & "'", cn, adOpenStatic, adLockReadOnly

If rec.BOF And rec.EOF Then
    ask = MsgBox("Do you really want to Cancel RO No. " & cmbrono & "?", vbYesNo, "Confirm Cancellation")
    If ask = vbYes Then
        cn.Execute "update romaster set cancelled=true where rono=" & Split(cmbrono, "/")(0) & " and finyear=" & Val("20" & Split(cmbrono, "/")(1))
        MsgBox "RO No. " & cmbrono & " has been Cancelled, which makes it unusable for both Printing and Billing.", vbInformation, "RO Cancelled !!"
        cmbrono = cmbrono
    End If
Else
    MsgBox "RO No. " & cmbrono & " is referred in Bill No. " & rec("billno") & "/" & Right(rec("finyear"), 2) & "/" & Right(rec("finyear") + 1, 2) & ".", vbCritical, "Not Allowed !!"
End If
Exit Sub
errorHandler:
MsgBox "Operation Failed!!"
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()
On Error GoTo errorHandler
Dim rec As New ADODB.Recordset
Dim Rep As New repro 'creating instance of crystal reports dsr
Dim frmrep As New frmreports 'new instance of report form

reprono = Split(cmbrono, "/")(0)
repfinyear = "20" & Split(cmbrono, "/")(1)

'**************************************Anish************************************************

'opening recordset for ro report
rec.Open "select RoMaster.*, RoDetail.* from RoMaster, RoDetail where RoMaster.RoMaster_ID=RoDetail.RoMaster_ID and RoMaster.Rono=" & reprono & " and romaster.finyear=" & repfinyear, cn, adOpenStatic, adLockReadOnly

'*******************************************************************************************

''rec.Open "SELECT RoMaster.*, RoDetail.* FROM RoDetail RIGHT JOIN RoMaster ON (RoDetail.finyear = RoMaster.finyear) AND (RoDetail.rono = RoMaster.rono) where romaster.rono=" & reprono & " and romaster.finyear=" & repfinyear, cn, adOpenStatic, adLockReadOnly

Rep.Database.SetDataSource rec  'assinging recordset to instance of cr dsr

frmrep.CRViewer1.ReportSource = Rep
frmrep.CRViewer1.PrintReport
frmrep.Caption = "Release Order"
Exit Sub
errorHandler:
MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdsave_Click()
 Dim rodate As Date, rec_rono As New ADODB.Recordset
 
 If cmdsave.Caption = "&Save" Then
 If IsValid = False Then Exit Sub
 
 rodate = Split(txtdate, "-")(1) & "/" & Split(txtdate, "-")(0) & "/" & Split(txtdate, "-")(2)
 rec_rono.Open "select max(rono) as maxno from romaster where finyear=" & finyear(rodate), cn, adOpenStatic, adLockReadOnly

 If (rec_rono.BOF And rec_rono.EOF) Or IsNull(rec_rono("maxno")) Then
    newro = 1
 Else
    newro = rec_rono("maxno") + 1
 End If
 
 On Error GoTo errorHandler
 
'******************************Anish*******************************************************
 Dim recMaxMasterID As New ADODB.Recordset  'fetching max id of romaster table
 Dim maxROMasID As Long                     'contains the max id
 
'saving in romaster table
 cn.Execute "insert into romaster (rono,finyear,rodate,client,admanager,material,language,cancelled) values ('" & newro & "','" & finyear(rodate) & "','" & rodate & "','" & Trim(txtclient) & "','" & Trim(txtadmanager) & "','" & Trim(txtmaterial) & "','" & Trim(txtlanguage) & "',false)"

'getting maximum romaster_Id from romaster table
 recMaxMasterID.Open "SELECT romaster_ID from romaster where rono=" & newro & " and finyear=" & finyear(rodate) & "", cn, adOpenForwardOnly, adLockReadOnly
    
'saving in rodetail
  For i = 0 To 2
    If Trim(txtcaption(i)) <> "" Then
        releasedate = CDate(Split(txtrdate(i), "-")(1) & "/" & Split(txtrdate(i), "-")(0) & "/" & Split(txtrdate(i), "-")(2))
        If Trim(cmbpremium(i)) <> "None" Or Trim(cmbpremium(i)) <> "" Then
            cn.Execute "insert into rodetail (Romaster_ID, caption, publication, edition, position, rate, releasedate, size1, size2, premium_name, premium_percent) values (" & recMaxMasterID("romaster_ID") & ",'" & Trim(txtcaption(i)) & "','" & Trim(cmbpublication(i)) & "','" & Trim(cmbedition(i)) & "','" & Trim(cmbposition(i)) & "','" & txtrate(i) & "','" & releasedate & "','" & Val(txtsize1(i)) & "','" & Val(txtsize2(i)) & "','" & Trim(cmbpremium(i)) & "','" & Val(txtppercent(i)) & "')"
        Else
            cn.Execute "insert into rodetail (Romaster_ID, caption, publication, edition, position, rate, releasedate, size1, size2, premium_name, premium_percent) values (" & recMaxMasterID("romaster_ID") & ",'" & Trim(txtcaption(i)) & "','" & Trim(cmbpublication(i)) & "','" & Trim(cmbedition(i)) & "','" & Trim(cmbposition(i)) & "','" & txtrate(i) & "','" & releasedate & "','" & Val(txtsize1(i)) & "','" & Val(txtsize2(i)) & "','',0)" 'set premium percent=0 and premium name = empty string
        End If
    End If
 Next
 
'*******************************************************************************************
'saving in master table
' cn.Execute "insert into romaster values ('" & newro & "','" & finyear(rodate) & "','" & rodate & "','" & Trim(txtclient) & "','" & Trim(txtadmanager) & "','" & Trim(txtmaterial) & "','" & Trim(txtlanguage) & "',false)"
'saving in detail table
' For i = 0 To 2
'    If Trim(txtcaption(i)) <> "" Then
'        releasedate = CDate(Split(txtrdate(i), "-")(1) & "/" & Split(txtrdate(i), "-")(0) & "/" & Split(txtrdate(i), "-")(2))
'        If Trim(cmbpremium(i)) <> "None" Or Trim(cmbpremium(i)) <> "" Then
'            cn.Execute "insert into rodetail values ('" & newro & "','" & finyear(rodate) & "','" & Trim(txtcaption(i)) & "','" & Trim(cmbpublication(i)) & "','" & Trim(cmbedition(i)) & "','" & Trim(cmbposition(i)) & "','" & txtrate(i) & "','" & releasedate & "','" & Val(txtsize1(i)) & "','" & Val(txtsize2(i)) & "','" & Trim(cmbpremium(i)) & "','" & Val(txtppercent(i)) & "')"
'        Else
'            cn.Execute "insert into rodetail values ('" & newro & "','" & finyear(rodate) & "','" & Trim(txtcaption(i)) & "','" & Trim(cmbpublication(i)) & "','" & Trim(cmbedition(i)) & "','" & Trim(cmbposition(i)) & "','" & txtrate(i) & "','" & releasedate & "','" & Val(txtsize1(i)) & "','" & Val(txtsize2(i)) & "','',0)" 'set premium percent=0 and premium name = empty string
'        End If
'    End If
' Next
 
 
 MsgBox "Release Order Saved into Database.", vbInformation
 
 addro
 cmbrono.ListIndex = cmbrono.ListCount - 1
 cmbrono.Enabled = True
 
 cmdprint.Enabled = True
 cmdcancel.Enabled = True
 cmdsave.Caption = "&New"
Exit Sub

errorHandler:

MsgBox "Error No. " & Err.Number & vbCrLf & Err.Description, vbCritical

'*********************************Anish*****************************************************
'deleting records in case of error
cn.Execute "delete * from romaster where rono=" & newro & " and finyear=" & finyear(rodate)
cn.Execute "delete * from rodetail where romaster_ID=" & recMaxMasterID("romaster_ID")

'*******************************************************************************************

''cn.Execute "delete * from romaster where rono=" & newro
''cn.Execute "delete * from rodetail where rono=" & newro

Else
cmbrono.ListIndex = -1
cmbrono.Enabled = False
ClearControls
txtdate = Format(Date, "dd-mm-yyyy")
addpublication
cmdsave.Caption = "&Save"
cmdprint.Enabled = False
cmdcancel.Enabled = False
End If

End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
ReDim arrRoId(0)
addro
If cmbrono.ListCount > 0 Then
cmbrono.ListIndex = cmbrono.ListCount - 1
Else
cmdsave.Value = True
End If
lbldate = "(Ex. " & Format(Date, "dd-mm-yyyy") & ")"
End Sub

Private Sub txtdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub 'allows backspace
If KeyAscii = 45 And (Len(txtdate) = 2 Or Len(txtdate) = 5) Then Exit Sub ' allows '-'
If KeyAscii <> 45 And (Len(txtdate) = 2 Or Len(txtdate) = 5) Then
KeyAscii = 0 ' do not allow anything except - in 3rd and 6th place
Beep
Exit Sub
End If

If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
Beep
End If


End Sub

Private Sub txtppercent_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub 'allows backspace
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txtrate_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub 'allows backspace
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txtrdate_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub 'allows backspace
If KeyAscii = 45 And (Len(txtrdate(Index)) = 2 Or Len(txtrdate(Index)) = 5) Then Exit Sub ' allows '-'
If KeyAscii <> 45 And (Len(txtrdate(Index)) = 2 Or Len(txtrdate(Index)) = 5) Then
KeyAscii = 0 ' do not allow anything except - in 3rd and 6th place
Beep
Exit Sub
End If

If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txtsize1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub 'allows backspace
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txtsize2_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub 'allows backspace
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
Beep
End If

End Sub
Sub addpublication()
Dim rec As New ADODB.Recordset

cmbpublication(0).Clear
cmbpublication(1).Clear
cmbpublication(2).Clear
Erase arrpubid
ReDim arrpubid(0)

'adding names in combo
rec.Open "select distinct(publication),name from publication,ratecard where publication=publication_id", cn, adOpenForwardOnly, adLockReadOnly

If rec.BOF = True And rec.EOF = True Then
    Exit Sub
End If

cmbpublication(0).AddItem rec!Name
cmbpublication(1).AddItem rec!Name
cmbpublication(2).AddItem rec!Name
arrpubid(0) = rec!publication
rec.MoveNext

While rec.EOF = False
    ReDim Preserve arrpubid(UBound(arrpubid) + 1)
    cmbpublication(0).AddItem rec!Name
    cmbpublication(1).AddItem rec!Name
    cmbpublication(2).AddItem rec!Name
    arrpubid(UBound(arrpubid)) = rec!publication
    rec.MoveNext
Wend

End Sub

Sub addEdition(Optional pubid As Long = 0, Optional idx As Byte = 0)
Dim rec As New ADODB.Recordset

cmbedition(idx).Clear

Select Case idx
Case Is = 0
    Erase arreditionid0
    ReDim arreditionid0(0)
Case Is = 1
    Erase arreditionid1
    ReDim arreditionid1(0)
Case Is = 2
    Erase arreditionid2
    ReDim arreditionid2(0)
End Select

'adding names in combo
rec.Open "select edition_id,name from ratecard,edition where edition=edition_id and publication=" & pubid, cn, adOpenForwardOnly, adLockReadOnly

If rec.BOF = True And rec.EOF = True Then
    Exit Sub
End If

Select Case idx
Case Is = 0
    cmbedition(idx).AddItem rec!Name
    arreditionid0(0) = rec!Edition_id
    rec.MoveNext
    While rec.EOF = False
    ReDim Preserve arreditionid0(UBound(arreditionid0) + 1)
    cmbedition(idx).AddItem rec!Name
    arreditionid0(UBound(arreditionid0)) = rec!Edition_id
    rec.MoveNext
    Wend
Case Is = 1
    cmbedition(idx).AddItem rec!Name
    arreditionid1(0) = rec!Edition_id
    rec.MoveNext
    While rec.EOF = False
    ReDim Preserve arreditionid1(UBound(arreditionid1) + 1)
    cmbedition(idx).AddItem rec!Name
    arreditionid1(UBound(arreditionid1)) = rec!Edition_id
    rec.MoveNext
    Wend
Case Is = 2
    cmbedition(idx).AddItem rec!Name
    arreditionid2(0) = rec!Edition_id
    rec.MoveNext
    While rec.EOF = False
    ReDim Preserve arreditionid2(UBound(arreditionid2) + 1)
    cmbedition(idx).AddItem rec!Name
    arreditionid2(UBound(arreditionid2)) = rec!Edition_id
    rec.MoveNext
    Wend
End Select
End Sub

Sub addPosition(Optional pubid As Long = 0, Optional editionid As Long = 0, Optional idx As Byte = 0)
Dim rec As New ADODB.Recordset
cmbposition(idx).Clear

Select Case idx
Case Is = 0
    Erase arrpositionid0
    ReDim arrpositionid0(0)
Case Is = 1
    Erase arrpositionid1
    ReDim arrpositionid1(0)
Case Is = 2
    Erase arrpositionid2
    ReDim arrpositionid2(0)
End Select

'adding names in combo

rec.Open "select position_id,name from ratecard,positions where positions=position_id and publication=" & pubid & " and edition=" & editionid, cn, adOpenForwardOnly, adLockReadOnly


If rec.BOF = True And rec.EOF = True Then
    Exit Sub
End If

Select Case idx
Case Is = 0
    cmbposition(idx).AddItem rec!Name
    arrpositionid0(0) = rec!position_id
    rec.MoveNext
    While rec.EOF = False
    ReDim Preserve arrpositionid0(UBound(arrpositionid0) + 1)
    cmbposition(idx).AddItem rec!Name
    arrpositionid0(UBound(arrpositionid0)) = rec!position_id
    rec.MoveNext
    Wend
Case Is = 1
    cmbposition(idx).AddItem rec!Name
    arrpositionid1(0) = rec!position_id
    rec.MoveNext
    While rec.EOF = False
    ReDim Preserve arrpositionid1(UBound(arrpositionid1) + 1)
    cmbposition(idx).AddItem rec!Name
    arrpositionid1(UBound(arrpositionid1)) = rec!position_id
    rec.MoveNext
    Wend
Case Is = 2
    cmbposition(idx).AddItem rec!Name
    arrpositionid2(0) = rec!position_id
    rec.MoveNext
    While rec.EOF = False
    ReDim Preserve arrpositionid2(UBound(arrpositionid2) + 1)
    cmbposition(idx).AddItem rec!Name
    arrpositionid2(UBound(arrpositionid2)) = rec!position_id
    rec.MoveNext
    Wend
End Select

End Sub


Sub addPremium(pubid As Long, idx As Byte)
Dim rec As New ADODB.Recordset

cmbpremium(idx).Clear
cmbpremium(idx).AddItem "None"

Select Case idx
Case Is = 0
    Erase arrpremiumid0
    ReDim arrpremiumid0(0)
Case Is = 1
    Erase arrpremiumid1
    ReDim arrpremiumid1(0)
Case Is = 2
    Erase arrpremiumid2
    ReDim arrpremiumid2(0)
End Select


'adding names in combo
rec.Open "select premium_id,premium_name from premium where publication=" & pubid, cn, adOpenForwardOnly, adLockReadOnly

If rec.BOF = True And rec.EOF = True Then
    Exit Sub
End If

Select Case idx
    Case Is = 0
    cmbpremium(idx).AddItem rec!premium_name
    arrpremiumid0(0) = rec!premium_id
    rec.MoveNext

    While rec.EOF = False
        ReDim Preserve arrpremiumid0(UBound(arrpremiumid0) + 1)
        cmbpremium(idx).AddItem rec!premium_name
        arrpremiumid0(UBound(arrpremiumid0)) = rec!premium_id
        rec.MoveNext
    Wend
    
    Case Is = 1
    cmbpremium(idx).AddItem rec!premium_name
    arrpremiumid1(0) = rec!premium_id
    rec.MoveNext

    While rec.EOF = False
        ReDim Preserve arrpremiumid1(UBound(arrpremiumid1) + 1)
        cmbpremium(idx).AddItem rec!premium_name
        arrpremiumid1(UBound(arrpremiumid1)) = rec!premium_id
        rec.MoveNext
    Wend
    
    Case Is = 2
    cmbpremium(idx).AddItem rec!premium_name
    arrpremiumid2(0) = rec!premium_id
    rec.MoveNext

    While rec.EOF = False
        ReDim Preserve arrpremiumid2(UBound(arrpremiumid2) + 1)
        cmbpremium(idx).AddItem rec!premium_name
        arrpremiumid2(UBound(arrpremiumid2)) = rec!premium_id
        rec.MoveNext
    Wend
End Select

End Sub
Function IsValid() As Boolean
 On Error GoTo errorHandler
 
 If IsDate(Split(txtdate, "-")(1) & "-" & Split(txtdate, "-")(0) & "-" & Split(txtdate, "-")(2)) = False Then
 IsValid = False
 MsgBox "Please enter a valid Date.", vbCritical, "Invalid Entry"
 txtdate.SetFocus
 Exit Function
 End If
 
 If Trim(txtclient) = "" Then
 MsgBox "Please enter a Client Name.", vbCritical, "Incomplete Entry"
 txtclient.SetFocus
 IsValid = False
 Exit Function
 End If
  
 If Trim(txtadmanager) = "" Then
 MsgBox "Please enter Advertising Manager.", vbCritical, "Incomplete Entry"
 txtadmanager.SetFocus
 IsValid = False
 Exit Function
 End If
 
 If Trim(txtcaption(0)) = "" Then
 MsgBox "Please enter Caption.", vbCritical, "Incomplete Entry"
 txtcaption(0).SetFocus
 IsValid = False
 Exit Function
 End If

 For i = 0 To 2
    If Trim(txtcaption(i)) <> "" Then
        If IsNumeric(txtrate(i)) = False Or Val(txtrate(i)) < 0 Or Val(txtrate(i)) > 10000 Then
            MsgBox "Please enter a valid Rate.", vbCritical, "Invalid Entry"
            txtrate(i).SetFocus
            IsValid = False
            Exit Function
        End If
        If Trim(txtppercent(i)) <> "" And (IsNumeric(txtppercent(i)) = False Or Val(txtppercent(i)) < 0) Then
            MsgBox "Please enter a valid Premium Percent.", vbCritical, "Invalid Entry"
            txtppercent(i).SetFocus
            IsValid = False
            Exit Function
        End If
        If IsDate(Split(txtrdate(i), "-")(1) & "-" & Split(txtrdate(i), "-")(0) & "-" & Split(txtrdate(i), "-")(2)) = False Then
            IsValid = False
            MsgBox "Please enter a valid Date.", vbCritical, "Invalid Entry"
            txtrdate(i).SetFocus
            Exit Function
        End If
        If IsNumeric(txtsize1(i)) = False Or Val(txtsize1(i)) < 0 Or Val(txtsize1(i)) > 10000 Then
            MsgBox "Please enter a valid Size1.", vbCritical, "Invalid Entry"
            txtsize1(i).SetFocus
            IsValid = False
            Exit Function
        End If
        If IsNumeric(txtsize2(i)) = False Or Val(txtsize2(i)) < 0 Or Val(txtsize2(i)) > 10000 Then
            MsgBox "Please enter a valid Size2.", vbCritical, "Invalid Entry"
            txtsize2(i).SetFocus
            IsValid = False
            Exit Function
        End If
    End If
Next
  
 If Trim(txtmaterial) = "" Then
 MsgBox "Please enter Material.", vbCritical, "Incomplete Entry"
 txtmaterial.SetFocus
 IsValid = False
 Exit Function
 End If
 
 If Trim(txtlanguage) = "" Then
 MsgBox "Please enter Language.", vbCritical, "Incomplete Entry"
 txtlanguage.SetFocus
 IsValid = False
 Exit Function
 End If
 
IsValid = True

Exit Function

errorHandler:
IsValid = False
MsgBox "Invalid Entry!!", vbCritical

End Function
Sub addro()
Dim rec As New ADODB.Recordset

cmbrono.Clear

'adding names in combo
rec.Open "select romaster_id,rono,finyear from romaster order by finyear,rono", cn, adOpenForwardOnly, adLockReadOnly

If rec.BOF = True And rec.EOF = True Then
    Exit Sub
End If

While rec.EOF = False
    temprono = rec("rono") & "/" & Right(rec("finyear"), 2) & "/" & Right(rec("finyear") + 1, 2)
    cmbrono.AddItem temprono
    arrRoId(UBound(arrRoId)) = rec("romaster_id")
    ReDim Preserve arrRoId(UBound(arrRoId) + 1)
    temprono = ""
    rec.MoveNext
Wend

End Sub

Sub ClearControls()
lblcancel.Visible = False
txtdate = ""
txtclient = ""
txtadmanager = ""
txtmaterial = ""
txtlanguage = ""
For i = 0 To 2
    cmbpublication(i) = ""
    cmbedition(i) = ""
    cmbposition(i) = ""
    cmbpremium(i) = ""
    txtcaption(i) = ""
    txtrdate(i) = ""
    txtrate(i) = ""
    txtppercent(i) = ""
    txtsize1(i) = ""
    txtsize2(i) = ""
Next
End Sub
