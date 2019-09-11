VERSION 5.00
Begin VB.Form FormPrincipal 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenação"
   ClientHeight    =   6975
   ClientLeft      =   1590
   ClientTop       =   1005
   ClientWidth     =   8565
   ForeColor       =   &H00404040&
   Icon            =   "FormOrdenacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8565
   Begin VB.CommandButton cmdAlgoritmos 
      BackColor       =   &H0000C000&
      Caption         =   "Algoritmos"
      Height          =   495
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   7560
      Top             =   3600
   End
   Begin VB.CommandButton cmdvoltar 
      BackColor       =   &H0000C000&
      Caption         =   "<<"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdavancar 
      BackColor       =   &H0000C000&
      Caption         =   ">>"
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdreiniciar 
      BackColor       =   &H0000C000&
      Caption         =   "Reiniciar"
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton CmdTerminei 
      BackColor       =   &H0000C000&
      Caption         =   "Terminei !"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdlimpar 
      BackColor       =   &H0000C000&
      Caption         =   "Limpar"
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Arrastar"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   4920
      Width           =   975
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      Height          =   1230
      Left            =   5160
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4560
      Width           =   3135
   End
   Begin VB.CommandButton Cmdcomparar 
      BackColor       =   &H0000C000&
      Caption         =   "Comparar"
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   13
      Left            =   7920
      TabIndex        =   71
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   12
      Left            =   7320
      TabIndex        =   70
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   11
      Left            =   6720
      TabIndex        =   69
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   10
      Left            =   6120
      TabIndex        =   68
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   9
      Left            =   5520
      TabIndex        =   67
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   8
      Left            =   4920
      TabIndex        =   66
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   7
      Left            =   4320
      TabIndex        =   65
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   6
      Left            =   3720
      TabIndex        =   64
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   63
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   62
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   61
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   60
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   59
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   58
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   13
      Left            =   7920
      TabIndex        =   57
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   12
      Left            =   7320
      TabIndex        =   56
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   11
      Left            =   6720
      TabIndex        =   55
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   10
      Left            =   6120
      TabIndex        =   54
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   9
      Left            =   5520
      TabIndex        =   53
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   8
      Left            =   4920
      TabIndex        =   52
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   7
      Left            =   4320
      TabIndex        =   51
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   6
      Left            =   3720
      TabIndex        =   50
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   49
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   48
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   47
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   46
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   45
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   44
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   7920
      TabIndex        =   43
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   7320
      TabIndex        =   42
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   6720
      TabIndex        =   41
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   6120
      TabIndex        =   40
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   5520
      TabIndex        =   39
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   4920
      TabIndex        =   38
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   13
      Left            =   7920
      TabIndex        =   37
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   7320
      TabIndex        =   36
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   11
      Left            =   6720
      TabIndex        =   35
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   10
      Left            =   6120
      TabIndex        =   34
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   5520
      TabIndex        =   33
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   4920
      TabIndex        =   32
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Tempo"
      Height          =   255
      Left            =   3120
      TabIndex        =   31
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label lblsegundos 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "0"
      Height          =   255
      Left            =   3960
      TabIndex        =   30
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label LblRespostas 
      BackStyle       =   0  'Transparent
      Caption         =   "Respostas :"
      Height          =   255
      Left            =   600
      TabIndex        =   29
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lbltentativas 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "0"
      Height          =   255
      Left            =   1920
      TabIndex        =   28
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Comparações feitas"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Maior"
      Height          =   255
      Left            =   8040
      TabIndex        =   26
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Menor"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   23
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1320
      TabIndex        =   22
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1920
      TabIndex        =   21
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2520
      TabIndex        =   20
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   3120
      TabIndex        =   19
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   3720
      TabIndex        =   18
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   4320
      TabIndex        =   17
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C000C0&
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   4320
      TabIndex        =   15
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   3720
      TabIndex        =   14
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   3120
      TabIndex        =   13
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2520
      TabIndex        =   12
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1920
      TabIndex        =   11
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1320
      TabIndex        =   10
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   9
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   495
   End
   Begin VB.Menu menger 
      Caption         =   "&Geral"
      Begin VB.Menu menrei 
         Caption         =   "Reiniciar"
         Shortcut        =   ^R
      End
      Begin VB.Menu mensai 
         Caption         =   "&Sair"
      End
   End
   Begin VB.Menu menpon 
      Caption         =   "&Ponutauções"
      Begin VB.Menu menVisualizar 
         Caption         =   "Visualizar"
         Shortcut        =   ^P
      End
      Begin VB.Menu menLimpar 
         Caption         =   "Limpar"
      End
   End
   Begin VB.Menu menocu 
      Caption         =   "&Mostrar"
      Begin VB.Menu menlet 
         Caption         =   "Letras"
      End
      Begin VB.Menu mencor 
         Caption         =   "Cores"
      End
      Begin VB.Menu menlcs 
         Caption         =   "Letras e cores"
         Checked         =   -1  'True
      End
      Begin VB.Menu menBarra 
         Caption         =   "-"
      End
      Begin VB.Menu menCom 
         Caption         =   "Comparações"
         Checked         =   -1  'True
      End
      Begin VB.Menu mentem 
         Caption         =   "Tempo"
         Checked         =   -1  'True
      End
      Begin VB.Menu menlis 
         Caption         =   "Lista"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu menaca 
      Caption         =   "Açã&o"
      Begin VB.Menu menpas 
         Caption         =   "Passar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mensol 
         Caption         =   "Soltar"
      End
   End
   Begin VB.Menu menPor 
      Caption         =   "&Comparar"
      Begin VB.Menu menPorLet 
         Caption         =   "Por letras"
      End
      Begin VB.Menu menPorCor 
         Caption         =   "Por cores"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu menalg 
      Caption         =   "Algoritmos"
      Begin VB.Menu menVerificarDetalhes 
         Caption         =   "Verificar detalhes"
      End
   End
   Begin VB.Menu menaju 
      Caption         =   "&Ajuda"
      Begin VB.Menu mentop 
         Caption         =   "Tópicos de ajuda"
         Shortcut        =   {F1}
      End
      Begin VB.Menu menResumo 
         Caption         =   "Resumo desta tela"
      End
      Begin VB.Menu mensob 
         Caption         =   "Sobre..."
      End
   End
End
Attribute VB_Name = "FormPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BooCompararPorLetra As Boolean
Dim dado(1 To 20) As String
Dim ContadorApertados As Byte
Dim SPrimeiro As String
Dim SSegundo As String
Dim Primeiro As Byte
Dim Segundo As Byte




Private Sub Check1_Click()
Dim cont As Byte
For cont = 0 To 13
    Label1(cont).DragMode = Check1.Value
Next cont
End Sub

Private Sub cmdAlgoritmos_Click()
   If jogoEmAndamento = True Then
      Exit Sub
   Else
      formComparacaoAlgoritmos.Show
   End If
End Sub

Private Sub cmdavancar_Click()
Dim cont As Integer
cont = 12
For cont = 12 To 0 Step -1
    If Label2(cont).BorderStyle = 1 And Label2(cont).BackColor <> &H8000000C Then
        Label2(cont + 1).BackColor = Label2(cont).BackColor
        Label2(cont + 1).Caption = Label2(cont).Caption
        Label2(cont + 1).ForeColor = Label2(cont).ForeColor
        Label2(cont).BackColor = &H8000000C
        Label2(cont).Caption = ""
        VetorUsuario(cont + 1, 0) = VetorUsuario(cont, 0)
        VetorUsuario(cont, 0) = 255
        VetorUsuario(cont + 1, 1) = VetorUsuario(cont, 1)
    End If
Next cont
End Sub

Private Sub Cmdcomparar_Click()
Dim cont As Byte
If ContadorApertados <> 2 Then
    MsgBox "Favor escolha duas figuras para comparação", vbInformation, "Ops..."
Else
    'Primeiro = 14
    SPrimeiro = ""
    For cont = 0 To 13
        If Label1(cont).BorderStyle = 1 Then
            If SPrimeiro = "" Then
                Primeiro = cont
                If BooCompararPorLetra = False Then
                    SPrimeiro = EquivalenteCor(cont)
                Else
                    SPrimeiro = EquivalenteLetra(cont)
                End If
            Else
                Segundo = cont
                If BooCompararPorLetra = False Then
                    SSegundo = EquivalenteCor(cont)
                Else
                    SSegundo = EquivalenteLetra(cont)
                End If
                cont = 13
            End If
        End If
    Next cont
       
   
    
    If Vetor(Primeiro) = Vetor(Segundo) Then
       MsgBox SPrimeiro & " = " & SSegundo, vbInformation, "Resultado"
        List1.AddItem (SPrimeiro & " = " & SSegundo)
    ElseIf Vetor(Primeiro) < Vetor(Segundo) Then
       MsgBox SPrimeiro & " < " & SSegundo, vbInformation, "Resultado"
        List1.AddItem (SPrimeiro & " < " & SSegundo)
    Else
       MsgBox SPrimeiro & " > " & SSegundo, vbInformation, "Resultado"
        List1.AddItem (SPrimeiro & " > " & SSegundo)
    End If
lbltentativas.Caption = lbltentativas.Caption + 1
    If lbltentativas.Caption < CInt(dado(6)) Then
        lbltentativas.BackColor = 49152
    ElseIf lbltentativas.Caption < CInt(dado(10)) Then
        lbltentativas.BackColor = &H80FFFF
    Else
        lbltentativas.BackColor = &HFF&
    End If
End If
End Sub

Private Sub cmdlimpar_Click()
Dim cont As Byte
    For cont = 0 To 13
        If Label2(cont).BorderStyle = 1 Then
            Label2(cont).BackColor = &H8000000C
            Label2(cont).Caption = ""
        End If
    Next cont
End Sub



Private Sub cmdreiniciar_Click()
    ReiniciarBotaoMenu
    Reiniciar

End Sub

Private Sub CmdTerminei_Click()
Dim cont As Byte
Dim acerto As Boolean
Dim repetido As Boolean
Dim recont As Byte
Dim branco As Boolean

repetido = False
acerto = True
branco = False


'RegistrarRecorde (InputBox(""))


cmdavancar.Enabled = False
cmdVoltar.Enabled = False
'LblRespostas.Visible = True

'-------------------------------
'    Teste
'-------------------------------
    'Dim tentativas As Integer
    'Dim tempo As Integer
'    lbltentativas.Caption = InputBox("Tentativas")
'    lblsegundos.Caption = InputBox("tempo")
'RegistrarRecorde ("teste")
'menpon_Click
    
'Exit Sub
'-------------------------------
'    Fim do teste
'-------------------------------


    For cont = 0 To 12 Step 1
        If VetorUsuario(cont, 0) > VetorUsuario(cont + 1, 0) Then acerto = False
        If VetorUsuario(cont, 0) = 255 Then
            branco = True
        Else
            Label4(cont).Caption = VetorUsuario(cont, 0)
        End If
        Label3(cont).Caption = Vetor(cont)
                             
        If cont <> 0 Then
            For recont = 0 To cont - 1
                If VetorUsuario(cont, 1) = VetorUsuario(recont, 1) Then repetido = True
            Next recont
        End If
    
        Label1(cont).Enabled = False
        Label2(cont).Enabled = False
        Label3(cont).Visible = True
        Label4(cont).Visible = True
                
    Next cont

Label1(13).Enabled = False ' Isso porque o cont acima foi só até 12 e fiz todas as comparações com cont até 12
Label2(13).Enabled = False ' Isso porque o cont acima foi só até 12 e fiz todas as comparações com cont até 12
CmdTerminei.Enabled = False
cmdlimpar.Enabled = False
Check1.Enabled = False
Cmdcomparar.Enabled = False

Label3(13).Caption = Vetor(13)

Label3(13).Visible = True
Label4(13).Visible = True

If VetorUsuario(13, 0) <> 255 Then
    Label4(13).Caption = VetorUsuario(13, 0)
Else
    branco = True
End If

If branco = True Then
    Timer1.Enabled = False
    MsgBox "Você deixou pelo menos uma das células da resposta em branco !", vbCritical, "Branco"
    Exit Sub
End If
    
    For recont = 0 To 12
        If VetorUsuario(recont, 1) = VetorUsuario(13, 1) Then repetido = True
    Next recont

If repetido = True Then
    Timer1.Enabled = False
    MsgBox "Você repetiu pelo menos uma cor na resposta", vbCritical, "Erro"
    Exit Sub
End If



If Vetor(13) = 255 Then acerto = False
If acerto = True Then
    Timer1.Enabled = False
    MsgBox "Ordem certa !", vbInformation, "Parabéns !"
    If lblsegundos.BackColor <> &HFF Or lbltentativas.BackColor <> &HFF Then
        Dim strrecorde
        strrecorde = "xxxxxxxxxx"
        While Len(strrecorde) > 9
            strrecorde = InputBox("Você obteve um desempenho fantástico. Digite o nome que você deseja que apareça no placar (USE DE 1 A 8 CARACTERES)", "Parabéns !")
        Wend
        If Trim(strrecorde) = "" Then strrecorde = "Anônimo"
        RegistrarRecorde (strrecorde)
    End If
    
    menVisualizar_Click
Else
    Timer1.Enabled = False
    MsgBox "Ordem errada !", vbCritical, "Erro !"
End If
    

End Sub

Private Sub cmdVoltar_Click()
Dim cont As Integer

For cont = 1 To 13 Step 1
    If Label2(cont).BorderStyle = 1 And Label2(cont).BackColor <> &H8000000C Then
        Label2(cont - 1).BackColor = Label2(cont).BackColor
        Label2(cont - 1).Caption = Label2(cont).Caption
        Label2(cont - 1).ForeColor = Label2(cont).ForeColor
        Label2(cont).Caption = ""
        Label2(cont).BackColor = &H8000000C
        VetorUsuario(cont - 1, 0) = VetorUsuario(cont, 0)
        VetorUsuario(cont, 0) = 255
        VetorUsuario(cont - 1, 1) = VetorUsuario(cont, 1)
    End If
Next cont

End Sub


Private Sub Form_initialize()
   Randomize Timer
   BooCompararPorLetra = False
   PreencheEquivalentes
   Reiniciar
   IniciarRecordes
End Sub


Private Sub Label1_Click(Index As Integer)
If Label1(Index).BorderStyle = 0 Then
    If ContadorApertados < 2 Then
        Label1(Index).BorderStyle = 1
        ContadorApertados = ContadorApertados + 1
    End If
Else
    Label1(Index).BorderStyle = 0
    ContadorApertados = ContadorApertados - 1
End If
End Sub

Private Sub Label2_Click(Index As Integer)
If Label2(Index).BorderStyle = 0 Then
        Label2(Index).BorderStyle = 1
Else
    Label2(Index).BorderStyle = 0
End If
End Sub

Sub PreencheEquivalentes()
    EquivalenteCor(0) = "Branco"
    EquivalenteCor(1) = "Vermelho"
    EquivalenteCor(2) = "Laranja"
    EquivalenteCor(3) = "Amarelo"
    EquivalenteCor(4) = "Verde claro"
    EquivalenteCor(5) = "Azul claro"
    EquivalenteCor(6) = "Azul escuro"
    EquivalenteCor(7) = "Lilás"
    EquivalenteCor(8) = "Verde escuro"
    EquivalenteCor(9) = "Cinza"
    EquivalenteCor(10) = "Preto"
    EquivalenteCor(11) = "Marrom"
    EquivalenteCor(12) = "Bege"
    EquivalenteCor(13) = "Azul marinho"
    EquivalenteLetra(0) = "A"
    EquivalenteLetra(1) = "B"
    EquivalenteLetra(2) = "C"
    EquivalenteLetra(3) = "D"
    EquivalenteLetra(4) = "E"
    EquivalenteLetra(5) = "F"
    EquivalenteLetra(6) = "G"
    EquivalenteLetra(7) = "H"
    EquivalenteLetra(8) = "I"
    EquivalenteLetra(9) = "J"
    EquivalenteLetra(10) = "K"
    EquivalenteLetra(11) = "L"
    EquivalenteLetra(12) = "M"
    EquivalenteLetra(13) = "N"
End Sub

Sub ConverteValor()
If BooCompararPorLetra = False Then
    If Primeiro = 0 Then
        SPrimeiro = "Branco"
    ElseIf Primeiro = 1 Then
        SPrimeiro = "Vermelho"
    ElseIf Primeiro = 2 Then
        SPrimeiro = "Laranja"
    ElseIf Primeiro = 3 Then
        SPrimeiro = "Amarelo"
    ElseIf Primeiro = 4 Then
        SPrimeiro = "Verde claro"
    ElseIf Primeiro = 5 Then
        SPrimeiro = "Azul claro"
    ElseIf Primeiro = 6 Then
        SPrimeiro = "Azul escuro"
    ElseIf Primeiro = 7 Then
        SPrimeiro = "Lilás"
    ElseIf Primeiro = 8 Then
        SPrimeiro = "Verde escuro"
    ElseIf Primeiro = 9 Then
        SPrimeiro = "Cinza"
    ElseIf Primeiro = 10 Then
        SPrimeiro = "Preto"
    ElseIf Primeiro = 11 Then
        SPrimeiro = "Marrom"
    ElseIf Primeiro = 12 Then
        SPrimeiro = "Bege"
    Else
        SPrimeiro = "Azul marinho"
    End If
    
    If Segundo = 0 Then
        SSegundo = "Branco"
    ElseIf Segundo = 1 Then
        SSegundo = "Vermelho"
    ElseIf Segundo = 2 Then
        SSegundo = "Laranja"
    ElseIf Segundo = 3 Then
        SSegundo = "Amarelo"
    ElseIf Segundo = 4 Then
        SSegundo = "Verde"
    ElseIf Segundo = 5 Then
        SSegundo = "Azul claro"
    ElseIf Segundo = 6 Then
        SSegundo = "Azul escuro"
    ElseIf Segundo = 7 Then
        SSegundo = "Lilás"
    ElseIf Segundo = 8 Then
        SSegundo = "Verde escuro"
    ElseIf Segundo = 9 Then
        SSegundo = "Cinza"
    ElseIf Segundo = 10 Then
        SSegundo = "Preto"
    ElseIf Segundo = 11 Then
        SSegundo = "Marrom"
    ElseIf Segundo = 12 Then
        SSegundo = "Bege"
    Else
        SSegundo = "Azul marinho"
    End If
End If

End Sub



Private Sub Label2_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
If mensol.Checked = True Then
    If Source.Index = 0 Or menlet.Checked = True Then
        Label2(Index).BackColor = &HFFFFFF
        Label2(Index).ForeColor = &H80000012
    ElseIf Source.Index = 1 Then
        Label2(Index).BackColor = &H8080FF
        Label2(Index).ForeColor = &H80000012
    ElseIf Source.Index = 2 Then
        Label2(Index).BackColor = &H80C0FF
        Label2(Index).ForeColor = &H80000012
    ElseIf Source.Index = 3 Then
        Label2(Index).BackColor = &HFFFF&
        Label2(Index).ForeColor = &H80000012
    ElseIf Source.Index = 4 Then
        Label2(Index).BackColor = &HC0FFC0
        Label2(Index).ForeColor = &H80000012
    ElseIf Source.Index = 5 Then
        Label2(Index).BackColor = &HFFFFC0
        Label2(Index).ForeColor = &H80000012
    ElseIf Source.Index = 6 Then
        Label2(Index).BackColor = &HFF0000
        Label2(Index).ForeColor = &H80000012
    ElseIf Source.Index = 7 Then
        Label2(Index).BackColor = &HC000C0
        Label2(Index).ForeColor = &H80000012
    ElseIf Source.Index = 8 Then
        Label2(Index).BackColor = &HC000&
        Label2(Index).ForeColor = &H80000012
    ElseIf Source.Index = 9 Then
        Label2(Index).BackColor = &HC0C0C0
        Label2(Index).ForeColor = &H80000012
    ElseIf Source.Index = 10 Then
        Label2(Index).BackColor = &H404000
        Label2(Index).ForeColor = &HFFFFFF
    ElseIf Source.Index = 11 Then
        Label2(Index).BackColor = &H4080&
        Label2(Index).ForeColor = &HFFFFFF
    ElseIf Source.Index = 12 Then
        Label2(Index).BackColor = &H80000018
        Label2(Index).ForeColor = &H80000012
    Else
       Label2(Index).BackColor = &H800000
       Label2(Index).ForeColor = &HFFFFFF
    End If
    Label2(Index).Caption = Source.Caption
    VetorUsuario(Index, 0) = Vetor(Source.Index)
    VetorUsuario(Index, 1) = Source.Index
End If
End Sub

Private Sub Label2_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
If menpas.Checked = True Then
    If State = 0 Then
        If Source.Index = 0 Or menlet.Checked = True Then
            Label2(Index).BackColor = &HFFFFFF
            Label2(Index).ForeColor = &H80000012
        ElseIf Source.Index = 1 Then
            Label2(Index).BackColor = &H8080FF
            Label2(Index).ForeColor = &H80000012
        ElseIf Source.Index = 2 Then
            Label2(Index).BackColor = &H80C0FF
            Label2(Index).ForeColor = &H80000012
        ElseIf Source.Index = 3 Then
            Label2(Index).BackColor = &HFFFF&
            Label2(Index).ForeColor = &H80000012
        ElseIf Source.Index = 4 Then
            Label2(Index).BackColor = &HC0FFC0
            Label2(Index).ForeColor = &H80000012
        ElseIf Source.Index = 5 Then
            Label2(Index).BackColor = &HFFFFC0
            Label2(Index).ForeColor = &H80000012
        ElseIf Source.Index = 6 Then
            Label2(Index).BackColor = &HFF0000
            Label2(Index).ForeColor = &H80000012
        ElseIf Source.Index = 7 Then
            Label2(Index).BackColor = &HC000C0
            Label2(Index).ForeColor = &H80000012
        ElseIf Source.Index = 8 Then
            Label2(Index).BackColor = &HC000&
            Label2(Index).ForeColor = &H80000012
        ElseIf Source.Index = 9 Then
            Label2(Index).BackColor = &HC0C0C0
            Label2(Index).ForeColor = &H80000012
        ElseIf Source.Index = 10 Then
            Label2(Index).BackColor = &H404000
            Label2(Index).ForeColor = &HFFFFFF
        ElseIf Source.Index = 11 Then
            Label2(Index).BackColor = &H4080&
            Label2(Index).ForeColor = &HFFFFFF
        ElseIf Source.Index = 12 Then
            Label2(Index).BackColor = &H80000018
            Label2(Index).ForeColor = &H80000012
        Else
           Label2(Index).BackColor = &H800000
           Label2(Index).ForeColor = &HFFFFFF
        End If
    
    Label2(Index).Caption = Source.Caption
    VetorUsuario(Index, 0) = Vetor(Source.Index)
    VetorUsuario(Index, 1) = Source.Index
    
    End If
End If
'If Source.Caption = "K" Or Source.Caption = "L" Or Source.Caption = "N" Then
'    Label2(Index).ForeColor = &HFFFFFF
'Else
'    Label2(Index).ForeColor = &H80000012
'End If


End Sub

Sub Reiniciar()
Dim cont As Byte
    For cont = 0 To Val(Right(Time, 2)) + Val(Mid(Time, 4, 2))
    Rnd
    Next cont
    For cont = 0 To 13
        Vetor(cont) = Int(Rnd * (253) + 1)
        VetorUsuario(cont, 0) = 255
    Next cont

cmdavancar.Enabled = True
cmdVoltar.Enabled = True

'LblRespostas.Visible = False
lblsegundos.Caption = 0
lblsegundos.BackColor = &HC000&
Timer1.Enabled = True
Check1.Value = 0
End Sub

Private Sub IniciarRecordes()

Dim caminho As String
Dim cont As Byte
caminho = App.Path & "\recordes2.rec"
If Dir(caminho) <> "" Then
    Open caminho For Input As #1
    For cont = 1 To 20
    If EOF(1) Then
        cont = 20
    Else
        Input #1, dado(cont)
        Debug.Print dado(cont)
    End If
    Next cont
    Close #1
Else
   Open App.Path & "\recordes2.rec" For Output As #2
   Print #2, "450" & vbCrLf & "550" & vbCrLf & "650" & vbCrLf & "750" & vbCrLf & "850" & vbCrLf & "30" & vbCrLf & "32" & vbCrLf & "35" & vbCrLf & "38" & vbCrLf & "40" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---"
        Close #2
        dado(1) = 450
        dado(2) = 550
        dado(3) = 650
        dado(4) = 750
        dado(5) = 850
        dado(6) = 30
        dado(7) = 32
        dado(8) = 35
        dado(9) = 38
        dado(10) = 40
        

            For cont = 11 To 20
                dado(cont) = "---"
            Next cont
End If
End Sub



Private Sub mansort_Click()
   If jogoEmAndamento = True Then Exit Sub
End Sub

Private Sub mencom_Click()
    menCom.Checked = Not (menCom.Checked)
    Label5.Visible = Not (Label5.Visible)
    lbltentativas.Visible = Not (lbltentativas.Visible)
End Sub

Private Sub mencor_Click()
If mencor.Checked = False Then
    Dim cont
    menPorLet.Enabled = False
    menPorLet.Checked = False
    menPorCor.Checked = True
    menPorCor.Enabled = True
    BooCompararPorLetra = False
    menlet.Checked = False
    mencor.Checked = True
    menlcs.Checked = False
    If Mid(List1.List(0), 3, 1) = ">" Or Mid(List1.List(0), 3, 1) = "<" Or Mid(List1.List(0), 3, 1) = "=" Then ConverterLista
    AparecerCor
    For cont = 0 To 13
        Label1(cont).Caption = ""
        Label2(cont).Caption = ""
    Next cont
End If
End Sub

Private Sub menlcs_Click()
If menlcs.Checked = False Then
    menPorCor.Enabled = True
    menPorLet.Enabled = True
    menlcs.Checked = True
    menlet.Checked = False
    mencor.Checked = False
    If Label1(1).BackColor <> &H8080FF Then
        AparecerCor
    Else
        AparecerLetra
    End If
End If
End Sub

Private Sub menlet_Click()
If menlet.Checked = False Then
    Dim cont As Byte
    menPorLet.Enabled = True
    menPorLet.Checked = True
    menPorCor.Checked = False
    menPorCor.Enabled = False
    BooCompararPorLetra = True
    menlet.Checked = True
    mencor.Checked = False
    menlcs.Checked = False
    If (Mid(List1.List(0), 3, 1) <> ">" Or Mid(List1.List(0), 3, 1) <> "<" Or Mid(List1.List(0), 3, 1) <> "=") And List1.ListCount <> 0 Then ConverterLista
    AparecerLetra
    For cont = 0 To 13
        Label1(cont).BackColor = &HFFFFFF
        Label1(cont).ForeColor = &H80000012
        If Label2(cont).BackColor <> &H8000000C Then Label2(cont).BackColor = &HFFFFFF
        'Label2(cont).BackColor = &HFFFFFF
        'Label2(cont).ForeColor = &H80000012
    Next cont
End If
End Sub

Private Sub menLimpar_Click()
   If MsgBox("Deseja realmente apagar as pontuações?", vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmação") = vbYes Then
    Open App.Path & "\recordes2.rec" For Output As #2
            Print #2, "450" & vbCrLf & "550" & vbCrLf & "650" & vbCrLf & "750" & vbCrLf & "850" & vbCrLf & "30" & vbCrLf & "32" & vbCrLf & "35" & vbCrLf & "38" & vbCrLf & "40" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---"
        Close #2
        dado(1) = 450
        dado(2) = 550
        dado(3) = 650
        dado(4) = 750
        dado(5) = 850
        dado(6) = 30
        dado(7) = 32
        dado(8) = 35
        dado(9) = 38
        dado(10) = 40
        
        Dim cont
            For cont = 11 To 20
                dado(cont) = "---"
            Next cont
    End If

End Sub

Private Sub menlis_Click()
    menlis.Checked = Not (menlis.Checked)
    List1.Visible = Not (List1.Visible)
End Sub

Private Sub menpas_Click()
    menpas.Checked = True
    mensol.Checked = False
End Sub

Private Sub menPorCor_Click()
    If menPorCor.Checked = False Then
        BooCompararPorLetra = False
        menPorCor.Checked = True
        menPorLet.Checked = False
        ConverterLista
    End If
End Sub

Private Sub menPorLet_Click()
    If menPorLet.Checked = False Then
        BooCompararPorLetra = True
        menPorCor.Checked = False
        menPorLet.Checked = True
        ConverterLista
    End If
End Sub

Private Sub menquick_Click()
    If jogoEmAndamento = True Then Exit Sub
End Sub

Private Sub menrei_Click()
    ReiniciarBotaoMenu
    Reiniciar
       
End Sub

Private Sub menResumo_Click()
   MsgBox "Esta é a tela principal, onde pode-se jogar, comparando as caixas da primeira fileria e arrastando-as fazendo uma fila em ordem crecente na segunda fileira. Uma descrição completa, com a função de cada botão, é encontrada no arquivo de ajuda", vbInformation, "Ajuda resumida"
End Sub

Private Sub mensai_Click()
    End
End Sub

Private Sub menShell_Click()
    If jogoEmAndamento = True Then Exit Sub
    formComparacaoAlgoritmos.cmbAlgoritmo.ItemData(4) = formComparacaoAlgoritmos.cmbAlgoritmo.List(4)
    formComparacaoAlgoritmos.Show 1
    
    formComparacaoAlgoritmos.shell
End Sub

Private Sub mensob_Click()
    formSobre.Show 1
End Sub

Private Sub mensol_Click()
    menpas.Checked = False
    mensol.Checked = True
End Sub

Private Sub mensort_Click()
    If jogoEmAndamento = True Then Exit Sub
    
End Sub

Private Sub mensort2_Click()
    If jogoEmAndamento = True Then Exit Sub
End Sub

Private Sub mentem_Click()
    mentem.Checked = Not (mentem.Checked)
    Label7.Visible = Not (Label7.Visible)
    lblsegundos.Visible = Not (lblsegundos.Visible)
End Sub

Private Sub mentop_Click()
 Dim strArquivo As String
 Dim objAjuda As ClasseAjuda
 
    Set objAjuda = New ClasseAjuda
    
    strArquivo = App.Path & "\ajuda\ajuda.chm"
    Call objAjuda.Show(strArquivo) ', "janelaHelp")
    Set objAjuda = Nothing
End Sub

Private Sub menVerificarDetalhes_Click()
   If jogoEmAndamento = True Then Exit Sub
   formComparacaoAlgoritmos.Show 1
End Sub

Private Sub menVisualizar_Click()
    MsgBox "As melhoeres pontuações são: " & vbCrLf & vbCrLf & "   Tempo:" & vbCrLf & "1º " & dado(11) & " - " & dado(1) & vbCrLf & "2º " & dado(12) & " - " & dado(2) & vbCrLf & "3º " & dado(13) & " - " & dado(3) & vbCrLf & "4º " & dado(14) & " - " & dado(4) & vbCrLf & "5º " & dado(15) & " - " & dado(5) & vbCrLf & vbCrLf & "   Comparações:" & vbCrLf & "1º " & dado(16) & " - " & dado(6) & vbCrLf & "2º " & dado(17) & " - " & dado(7) & vbCrLf & "3º " & dado(18) & " - " & dado(8) & vbCrLf & "4º " & dado(19) & " - " & dado(9) & vbCrLf & "5º " & dado(20) & " - " & dado(10), vbInformation, "Recordes"
End Sub

Private Sub Timer1_Timer()
    lblsegundos.Caption = lblsegundos.Caption + 1
    If lblsegundos.Caption < CDbl(dado(1)) Then
        lblsegundos.BackColor = &HC000&
    ElseIf lblsegundos.Caption < CDbl(dado(5)) Then
        lblsegundos.BackColor = &H80FFFF
    Else
        lblsegundos.BackColor = &HFF&
    End If
End Sub

Sub RegistrarRecorde(strecorde As String)

'Recorde de tempo
    If lblsegundos.Caption < CDbl(dado(3)) Then
        If lblsegundos.Caption < CDbl(dado(1)) Then
            dado(5) = dado(4)
            dado(15) = dado(14)
            dado(4) = dado(3)
            dado(14) = dado(13)
            dado(3) = dado(2)
            dado(13) = dado(12)
            dado(2) = dado(1)
            dado(12) = dado(11)
            dado(1) = lblsegundos.Caption
            dado(11) = strecorde
        ElseIf lblsegundos.Caption < CDbl(dado(2)) Then
            dado(5) = dado(4)
            dado(15) = dado(14)
            dado(4) = dado(3)
            dado(14) = dado(13)
            dado(3) = dado(2)
            dado(13) = dado(12)
            dado(2) = lblsegundos.Caption
            dado(12) = strecorde
        Else
            dado(5) = dado(4)
            dado(15) = dado(14)
            dado(4) = dado(3)
            dado(14) = dado(13)
            dado(3) = lblsegundos.Caption
            dado(13) = strecorde
        End If
    ElseIf lblsegundos.Caption < CDbl(dado(5)) Then
        If lblsegundos.Caption < CDbl(dado(4)) Then
            dado(5) = dado(4)
            dado(15) = dado(14)
            dado(4) = lblsegundos.Caption
            dado(14) = strecorde
        Else
            dado(5) = lblsegundos.Caption
            dado(15) = strecorde
        End If
    End If
    
'Recorde de comparações
    If lbltentativas.Caption < CDbl(dado(8)) Then
        If lbltentativas.Caption < CDbl(dado(6)) Then
            dado(10) = dado(9)
            dado(20) = dado(19)
            dado(9) = dado(8)
            dado(19) = dado(18)
            dado(8) = dado(7)
            dado(18) = dado(17)
            dado(7) = dado(6)
            dado(17) = dado(16)
            dado(6) = lbltentativas.Caption
            dado(16) = strecorde
        ElseIf lbltentativas.Caption < CDbl(dado(7)) Then
            dado(10) = dado(9)
            dado(20) = dado(19)
            dado(9) = dado(8)
            dado(19) = dado(18)
            dado(8) = dado(7)
            dado(18) = dado(17)
            dado(7) = lbltentativas.Caption
            dado(17) = strecorde
        Else
            dado(10) = dado(9)
            dado(20) = dado(19)
            dado(9) = dado(8)
            dado(19) = dado(18)
            dado(8) = lbltentativas.Caption
            dado(18) = strecorde
        End If
    ElseIf lbltentativas.Caption < CDbl(dado(10)) Then
        If lbltentativas.Caption < CDbl(dado(9)) Then
            dado(10) = dado(9)
            dado(20) = dado(19)
            dado(9) = lbltentativas.Caption
            dado(19) = strecorde
        Else
            dado(10) = lbltentativas.Caption
            dado(20) = strecorde
        End If
    End If
    
Open App.Path & "\recordes2.rec" For Output As #3
Dim cont As Byte
    For cont = 1 To 20
        Print #3, dado(cont)
    Next cont
Close #3
End Sub

Sub ReiniciarBotaoMenu()

Dim cont As Byte
    For cont = 0 To 13
        Label3(cont).Visible = False
        Label4(cont).Visible = False
        Label3(cont).Caption = ""
        Label4(cont).Caption = ""
        Label1(cont).Enabled = True
        Label2(cont).Enabled = True
        Label2(cont).Caption = ""
        Label1(cont).BorderStyle = 0
        Label2(cont).BorderStyle = 0
        Label2(cont).BackColor = &H8000000C
        lbltentativas.Caption = "0"
        VetorUsuario(cont, 1) = 0
    Next cont
    List1.Clear
    CmdTerminei.Enabled = True
    cmdlimpar.Enabled = True
    Check1.Enabled = True
    Cmdcomparar.Enabled = True
    ContadorApertados = 0
End Sub



Sub ConverterLista()
Dim cont As Byte
Dim recont As Byte
Dim valorconversao As String

If List1.ListCount = 0 Then Exit Sub

    If Mid(List1.List(0), 3, 1) = ">" Or Mid(List1.List(0), 3, 1) = "<" Or Mid(List1.List(0), 3, 1) = "=" Then
        For cont = 0 To List1.ListCount - 1
            For recont = 1 To 5 Step 4
                If Mid(List1.List(cont), recont, 1) = "A" Then
                    valorconversao = valorconversao & "Branco"
                ElseIf Mid(List1.List(cont), recont, 1) = "B" Then
                    valorconversao = valorconversao & "Vermelho"
                ElseIf Mid(List1.List(cont), recont, 1) = "C" Then
                    valorconversao = valorconversao & "Laranja"
                ElseIf Mid(List1.List(cont), recont, 1) = "D" Then
                    valorconversao = valorconversao & "Amarelo"
                ElseIf Mid(List1.List(cont), recont, 1) = "E" Then
                    valorconversao = valorconversao & "Verde claro"
                ElseIf Mid(List1.List(cont), recont, 1) = "F" Then
                    valorconversao = valorconversao & "Azul claro"
                ElseIf Mid(List1.List(cont), recont, 1) = "G" Then
                    valorconversao = valorconversao & "Azul escuro"
                ElseIf Mid(List1.List(cont), recont, 1) = "H" Then
                    valorconversao = valorconversao & "Lilás"
                ElseIf Mid(List1.List(cont), recont, 1) = "I" Then
                    valorconversao = valorconversao & "Verde escuro"
                ElseIf Mid(List1.List(cont), recont, 1) = "J" Then
                    valorconversao = valorconversao & "Cinza"
                ElseIf Mid(List1.List(cont), recont, 1) = "K" Then
                    valorconversao = valorconversao & "Preto"
                ElseIf Mid(List1.List(cont), recont, 1) = "L" Then
                    valorconversao = valorconversao & "Marrom"
                ElseIf Mid(List1.List(cont), recont, 1) = "M" Then
                    valorconversao = valorconversao & "Bege"
                Else
                    valorconversao = valorconversao & "Azul marinho"
                End If
                If recont = 1 Then valorconversao = valorconversao & Mid(List1.List(cont), 2, 3)
            Next recont
        List1.List(cont) = valorconversao
        valorconversao = ""
        Next cont
    Else
        Dim FinalCont As Byte
        Dim compfinal As Byte
        For cont = 0 To List1.ListCount - 1
            
            FinalCont = InStr(1, List1.List(cont), ">")
            If FinalCont = 0 Then
                FinalCont = InStr(1, List1.List(cont), "<")
                If FinalCont = 0 Then FinalCont = InStr(1, List1.List(cont), "=")
            End If
            
            For recont = 1 To FinalCont + 2 Step FinalCont + 1
                If recont <> 1 Then
                    valorconversao = valorconversao & Mid(List1.List(cont), recont - 3, 3)
                    compfinal = Len(List1.List(cont))
                Else
                    compfinal = FinalCont - 2
                End If
                
                If Mid(List1.List(cont), recont, compfinal) = "Branco" Then
                    valorconversao = valorconversao & "A"
                ElseIf Mid(List1.List(cont), recont, compfinal) = "Vermelho" Then
                    valorconversao = valorconversao & "B"
                ElseIf Mid(List1.List(cont), recont, compfinal) = "Laranja" Then
                    valorconversao = valorconversao & "C"
                ElseIf Mid(List1.List(cont), recont, compfinal) = "Amarelo" Then
                    valorconversao = valorconversao & "D"
                ElseIf Mid(List1.List(cont), recont, compfinal) = "Verde claro" Then
                    valorconversao = valorconversao & "E"
                ElseIf Mid(List1.List(cont), recont, compfinal) = "Azul claro" Then
                    valorconversao = valorconversao & "F"
                ElseIf Mid(List1.List(cont), recont, compfinal) = "Azul escuro" Then
                    valorconversao = valorconversao & "G"
                ElseIf Mid(List1.List(cont), recont, compfinal) = "Lilás" Then
                    valorconversao = valorconversao & "H"
                ElseIf Mid(List1.List(cont), recont, compfinal) = "Verde escuro" Then
                    valorconversao = valorconversao & "I"
                ElseIf Mid(List1.List(cont), recont, compfinal) = "Cinza" Then
                    valorconversao = valorconversao & "J"
                ElseIf Mid(List1.List(cont), recont, compfinal) = "Preto" Then
                    valorconversao = valorconversao & "K"
                ElseIf Mid(List1.List(cont), recont, compfinal) = "Marrom" Then
                    valorconversao = valorconversao & "L"
                ElseIf Mid(List1.List(cont), recont, compfinal) = "Bege" Then
                    valorconversao = valorconversao & "M"
                Else
                    valorconversao = valorconversao & "N"
                End If
                
            Next recont
        List1.List(cont) = valorconversao
        valorconversao = ""
        Next cont
    End If
End Sub

Sub AparecerCor()
    Dim cont
    Label1(0).BackColor = &HFFFFFF
    Label1(1).BackColor = &H8080FF
    Label1(2).BackColor = &H80C0FF
    Label1(3).BackColor = &HFFFF&
    Label1(4).BackColor = &HC0FFC0
    Label1(5).BackColor = &HFFFFC0
    Label1(6).BackColor = &HFF0000
    Label1(7).BackColor = &HC000C0
    Label1(8).BackColor = &HC000&
    Label1(9).BackColor = &HC0C0C0
    Label1(10).BackColor = &H404000
    Label1(10).ForeColor = &HFFFFFF
    Label1(11).BackColor = &H4080&
    Label1(11).ForeColor = &HFFFFFF
    Label1(12).BackColor = &H80000018
    Label1(13).BackColor = &H800000
    Label1(13).ForeColor = &HFFFFFF
    
    For cont = 0 To 13
        If Label2(cont).Caption = "A" Then
            Label2(cont).BackColor = &HFFFFFF
        ElseIf Label2(cont).Caption = "B" Then
            Label2(cont).BackColor = &H8080FF
        ElseIf Label2(cont).Caption = "C" Then
            Label2(cont).BackColor = &H80C0FF
        ElseIf Label2(cont).Caption = "D" Then
            Label2(cont).BackColor = &HFFFF&
        ElseIf Label2(cont).Caption = "E" Then
            Label2(cont).BackColor = &HC0FFC0
        ElseIf Label2(cont).Caption = "F" Then
            Label2(cont).BackColor = &HFFFFC0
        ElseIf Label2(cont).Caption = "G" Then
            Label2(cont).BackColor = &HFF0000
        ElseIf Label2(cont).Caption = "H" Then
            Label2(cont).BackColor = &HC000C0
        ElseIf Label2(cont).Caption = "I" Then
            Label2(cont).BackColor = &HC000&
        ElseIf Label2(cont).Caption = "J" Then
            Label2(cont).BackColor = &HC0C0C0
        ElseIf Label2(cont).Caption = "K" Then
            Label2(cont).BackColor = &H404000
        ElseIf Label2(cont).Caption = "L" Then
            Label2(cont).BackColor = &H4080&
        ElseIf Label2(cont).Caption = "M" Then
            Label2(cont).BackColor = &H80000018
        ElseIf Label2(cont).Caption = "N" Then
            Label2(cont).BackColor = &H800000
        End If
    Next cont
End Sub

Sub AparecerLetra()
    Label1(0).Caption = "A"
    Label1(1).Caption = "B"
    Label1(2).Caption = "C"
    Label1(3).Caption = "D"
    Label1(4).Caption = "E"
    Label1(5).Caption = "F"
    Label1(6).Caption = "G"
    Label1(7).Caption = "H"
    Label1(8).Caption = "I"
    Label1(9).Caption = "J"
    Label1(10).Caption = "K"
    Label1(11).Caption = "L"
    Label1(12).Caption = "M"
    Label1(13).Caption = "N"
    
    
    Dim cont
    For cont = 0 To 13
        If Label2(cont).BackColor = &HFFFFFF Then
            Label2(cont).Caption = "A"
        ElseIf Label2(cont).BackColor = &H8080FF Then
            Label2(cont).Caption = "B"
        ElseIf Label2(cont).BackColor = &H80C0FF Then
            Label2(cont).Caption = "C"
        ElseIf Label2(cont).BackColor = &HFFFF& Then
            Label2(cont).Caption = "D"
        ElseIf Label2(cont).BackColor = &HC0FFC0 Then
            Label2(cont).Caption = "E"
        ElseIf Label2(cont).BackColor = &HFFFFC0 Then
            Label2(cont).Caption = "F"
        ElseIf Label2(cont).BackColor = &HFF0000 Then
            Label2(cont).Caption = "G"
        ElseIf Label2(cont).BackColor = &HC000C0 Then
            Label2(cont).Caption = "H"
        ElseIf Label2(cont).BackColor = &HC000& Then
            Label2(cont).Caption = "I"
        ElseIf Label2(cont).BackColor = &HC0C0C0 Then
            Label2(cont).Caption = "J"
        ElseIf Label2(cont).BackColor = &H404000 Then
            Label2(cont).Caption = "K"
        ElseIf Label2(cont).BackColor = &H4080& Then
            Label2(cont).Caption = "L"
        ElseIf Label2(cont).BackColor = &H80000018 Then
            Label2(cont).Caption = "M"
        ElseIf Label2(cont).BackColor = &H800000 Then
            Label2(cont).Caption = "N"
        End If
    Next cont
End Sub

Function jogoEmAndamento() As Boolean
   If cmdlimpar.Enabled = True Then
      MsgBox "O jogo deve ser encerrado, utilizando o botão 'Terminei !', antes de verificar os algoritmos e suas resoluções", vbInformation, "Necessário encerramento"
   Else

   End If
   jogoEmAndamento = cmdlimpar.Enabled
End Function
