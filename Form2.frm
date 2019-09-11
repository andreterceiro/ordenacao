VERSION 5.00
Begin VB.Form formSobre 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sobre ..."
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdvoltar 
      BackColor       =   &H0000C000&
      Caption         =   "Voltar"
      Height          =   495
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   975
   End
   Begin VB.Line Line9 
      X1              =   3480
      X2              =   3360
      Y1              =   1320
      Y2              =   1440
   End
   Begin VB.Line Line8 
      X1              =   4080
      X2              =   3960
      Y1              =   1320
      Y2              =   1440
   End
   Begin VB.Line Line7 
      X1              =   3480
      X2              =   3360
      Y1              =   840
      Y2              =   960
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3480
      Top             =   840
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   3360
      Top             =   960
      Width           =   615
   End
   Begin VB.Line Line6 
      X1              =   2040
      X2              =   1920
      Y1              =   840
      Y2              =   960
   End
   Begin VB.Line Line5 
      X1              =   2040
      X2              =   1920
      Y1              =   1320
      Y2              =   1440
   End
   Begin VB.Line Line4 
      X1              =   2640
      X2              =   2520
      Y1              =   1320
      Y2              =   1440
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2040
      Top             =   840
      Width           =   615
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   1920
      Top             =   960
      Width           =   615
   End
   Begin VB.Line Line3 
      X1              =   600
      X2              =   480
      Y1              =   1320
      Y2              =   1440
   End
   Begin VB.Line Line2 
      X1              =   1200
      X2              =   1080
      Y1              =   1320
      Y2              =   1440
   End
   Begin VB.Line Line1 
      X1              =   600
      X2              =   480
      Y1              =   840
      Y2              =   960
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   600
      Top             =   840
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   480
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "andreterceiro@yahoo.com.br"
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
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Desenvolvido por André de Paula Terceiro"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ordenação V 2.0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "formSobre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

