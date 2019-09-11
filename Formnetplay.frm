VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Formnetplay 
   Caption         =   "Netplay"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   LinkTopic       =   "Form4"
   ScaleHeight     =   3060
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optcliente 
      Caption         =   "Atuar como cliente"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   360
      Width           =   1695
   End
   Begin VB.OptionButton optservidor 
      Caption         =   "Atuar como servidor"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ok"
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdcancelar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1800
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtporta 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtip 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Porta"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "IP"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "Formnetplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
