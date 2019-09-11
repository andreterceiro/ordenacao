VERSION 5.00
Begin VB.Form FormAlgoritmos 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   6510
   ClientLeft      =   1590
   ClientTop       =   1395
   ClientWidth     =   9195
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   9195
   Begin VB.CommandButton cmdn2 
      BackColor       =   &H0000C000&
      Caption         =   "OrdV2"
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton cmdquick 
      BackColor       =   &H0000C000&
      Caption         =   "Quick sort"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5760
      Width           =   975
   End
   Begin VB.ListBox lstn2 
      Height          =   840
      Left            =   5160
      TabIndex        =   13
      Top             =   4560
      Width           =   3735
   End
   Begin VB.ListBox lstquick 
      Height          =   840
      Left            =   240
      TabIndex        =   12
      Top             =   4560
      Width           =   3735
   End
   Begin VB.CommandButton cmdshell 
      BackColor       =   &H0000C000&
      Caption         =   "Shell sort"
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdsort2 
      BackColor       =   &H0000C000&
      Caption         =   "Sort 2"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   975
   End
   Begin VB.ListBox lstshell 
      Height          =   840
      Left            =   5160
      TabIndex        =   7
      Top             =   2640
      Width           =   3735
   End
   Begin VB.ListBox lstsort2 
      Height          =   840
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   3735
   End
   Begin VB.CommandButton cmdbolha 
      BackColor       =   &H0000C000&
      Caption         =   "Bolha"
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdsort 
      BackColor       =   &H0000C000&
      Caption         =   "Sort"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.ListBox lstbolha 
      Height          =   840
      Left            =   5160
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.ListBox lstsort 
      Height          =   840
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label lbln2 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   7560
      TabIndex        =   15
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label lblquick 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   2640
      TabIndex        =   14
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label lblshell 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   7560
      TabIndex        =   9
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label lblsort2 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label lblbolha 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   7560
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblsort 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Menu menalg 
      Caption         =   "&Algorítimos"
      Begin VB.Menu mensort 
         Caption         =   "Sort"
      End
      Begin VB.Menu menbolha 
         Caption         =   "Bolha"
      End
      Begin VB.Menu mensort2 
         Caption         =   "Sort 2"
      End
      Begin VB.Menu menshellsort 
         Caption         =   "Shell sort"
      End
      Begin VB.Menu menquicksort 
         Caption         =   "Quick sort"
      End
      Begin VB.Menu menn2 
         Caption         =   "N 2"
      End
   End
End
Attribute VB_Name = "FormAlgoritmos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m As Byte
Dim X As Byte
Dim s  As Integer
Dim f As Integer
Dim i As Byte
Dim j As Byte
Dim l As Integer
Dim teste2 As Boolean
Dim vetorusuario1(1 To 14) As Byte
Dim comparacoes As Byte
Dim aux As Integer
Dim f1(14) As Integer
Dim l1(14) As Integer

Private Sub cmdbolha_Click()
    bolha
End Sub

Private Sub cmdn2_Click()
    binn2
End Sub

Private Sub cmdquick_Click()
    quick
End Sub

Private Sub cmdshell_Click()
shell
End Sub

Private Sub cmdsort_Click()
    sort
End Sub

Private Sub cmdsort2_Click()
    sort2
End Sub

Sub sort()
Dim cont As Integer
Dim aux As Byte
Dim comparacoes As Integer
'Dim ok As Integer

For cont = 0 To 13
    VetorUsuario(cont, 0) = Vetor(cont) ' Preenchendo o vetor usuario com os valores do vetor
    VetorUsuario(cont, 1) = cont ' necessário para preencher o listbox corretamente
Next cont
cont = 0


For cont = 0 To 12
    comparacoes = comparacoes + 1
    If VetorUsuario(cont, 0) > VetorUsuario(cont + 1, 0) Then
        'ok = ok + 1
        aux = VetorUsuario(cont, 0)
        VetorUsuario(cont, 0) = VetorUsuario(cont + 1, 0)
        VetorUsuario(cont + 1, 0) = aux
        If FormPrincipal.menPorLet.Checked = True Then
            lstsort.AddItem (EquivalenteLetra(VetorUsuario(cont, 1)) & " > " & EquivalenteLetra(VetorUsuario(cont + 1, 1))) & " - Trocou e reiniciou"
        Else
            lstsort.AddItem (EquivalenteCor(VetorUsuario(cont, 1)) & " > " & EquivalenteCor(VetorUsuario(cont + 1, 1))) & " - Trocou e reiniciou"
        End If
        aux = VetorUsuario(cont + 1, 1)
        VetorUsuario(cont + 1, 1) = VetorUsuario(cont, 1)
        VetorUsuario(cont, 1) = aux
        cont = -1
    Else
        If FormPrincipal.menPorLet.Checked = True Then
            lstsort.AddItem (EquivalenteLetra(VetorUsuario(cont, 1)) & " <= " & EquivalenteLetra(VetorUsuario(cont + 1, 1)))
        Else
            lstsort.AddItem (EquivalenteCor(VetorUsuario(cont, 1)) & " <= " & EquivalenteCor(VetorUsuario(cont + 1, 1)))
        End If
    End If
Next cont
MsgBox "comaparacoes = " & comparacoes, vbInformation, "Ok !"
lblsort.Caption = comparacoes & " comparações"
End Sub

Sub bolha()
Dim cont As Integer
Dim recont As Byte
Dim aux As Byte
Dim comparacoes As Integer
'Dim ok As Integer
Dim eq As Byte 'Para acertar o preenchimento do listbox

For cont = 0 To 13
    VetorUsuario(cont, 0) = Vetor(cont) ' Preenchendo o vetor usuario com os valores do vetor
Next cont
cont = 0

For cont = 0 To 12
    eq = cont
    For recont = (cont + 1) To (13)
        comparacoes = comparacoes + 1
        If VetorUsuario(cont, 0) > VetorUsuario(recont, 0) Then
            'ok = ok + 1
            aux = VetorUsuario(recont, 0)
            VetorUsuario(recont, 0) = VetorUsuario(cont, 0)
            VetorUsuario(cont, 0) = aux
            If FormPrincipal.menPorLet.Checked = True Then
                lstbolha.AddItem (EquivalenteLetra(eq) & " > " & EquivalenteLetra(recont)) & " - Trocou os 2"
            Else
                lstbolha.AddItem (EquivalenteCor(eq) & " > " & EquivalenteCor(recont)) & " - Trocou os 2"
            End If
            
         eq = recont
         Else
            If FormPrincipal.menPorLet.Checked = True Then
                lstbolha.AddItem (EquivalenteLetra(eq) & " <= " & EquivalenteLetra(recont))
            Else
                lstbolha.AddItem (EquivalenteCor(eq) & " <= " & EquivalenteCor(recont))
            End If
        End If
    Next recont
lstbolha.List(lstbolha.ListCount - 1) = lstbolha.List(lstbolha.ListCount - 1) & " - Achou o" & cont + 1 & "º"
Next cont
MsgBox "comaparacoes = " & comparacoes, vbInformation, "Ok !"
lblbolha.Caption = comparacoes & " comparações"
End Sub

Sub sort2()
Dim cont As Integer
Dim recont As Integer
Dim aux As Byte
Dim maxcont As Byte
Dim comparacoes As Integer
'Dim ok As Integer

For cont = 0 To 13
    VetorUsuario(cont, 0) = Vetor(cont) ' Preenchendo o vetor usuario com os valores do vetor
    VetorUsuario(cont, 1) = cont
Next cont
cont = 0


For cont = 0 To 12
    comparacoes = comparacoes + 1
    If VetorUsuario(cont, 0) > VetorUsuario(cont + 1, 0) Then
 '       ok = ok + 1
        aux = VetorUsuario(cont, 0)
        VetorUsuario(cont, 0) = VetorUsuario(cont + 1, 0)
        VetorUsuario(cont + 1, 0) = aux
        
        aux = VetorUsuario(cont, 1)
        VetorUsuario(cont, 1) = VetorUsuario(cont + 1, 1)
        VetorUsuario(cont + 1, 1) = aux
        
        If FormPrincipal.menPorLet.Checked = True Then
            lstsort2.AddItem (EquivalenteLetra(VetorUsuario(cont + 1, 1)) & " > " & EquivalenteLetra(VetorUsuario(cont, 1))) & " - Troca os 2"
        Else
            lstsort2.AddItem (EquivalenteCor(VetorUsuario(cont + 1, 1)) & " > " & EquivalenteCor(VetorUsuario(cont, 1))) & " - Troca os 2"
        End If
        
        If cont <> 0 Then
            lstsort2.List(lstsort2.ListCount - 1) = lstsort2.List(lstsort2.ListCount - 1) & " e compara com os anteriores"
            For recont = cont To 1 Step -1
                comparacoes = comparacoes + 1
                If VetorUsuario(recont - 1, 0) > VetorUsuario(recont, 0) Then
  '                  ok = ok + 1
                    aux = VetorUsuario(recont, 0)
                    VetorUsuario(recont, 0) = VetorUsuario(recont - 1, 0)
                    VetorUsuario(recont - 1, 0) = aux
                    
                    aux = VetorUsuario(recont, 1)
                    VetorUsuario(recont, 1) = VetorUsuario(recont - 1, 1)
                    VetorUsuario(recont - 1, 1) = aux
                    
                    If FormPrincipal.menPorLet.Checked = True Then
                        lstsort2.AddItem (EquivalenteLetra(VetorUsuario(recont, 1)) & " > " & EquivalenteLetra(VetorUsuario(recont - 1, 1))) & " - Troca os 2"
                    Else
                        lstsort2.AddItem (EquivalenteCor(VetorUsuario(recont, 1)) & " > " & EquivalenteCor(VetorUsuario(recont - 1, 1))) & " - Troca os 2"
                    End If
                 Else
                    If FormPrincipal.menPorLet.Checked = True Then
                        lstsort2.AddItem (EquivalenteLetra(VetorUsuario(recont - 1, 1)) & " <= " & EquivalenteLetra(VetorUsuario(recont, 1)))
                    Else
                        lstsort2.AddItem (EquivalenteCor(VetorUsuario(recont - 1, 1)) & " <= " & EquivalenteCor(VetorUsuario(recont, 1)))
                    End If
                    recont = 1
                End If
            Next recont
        End If
    Else
        If FormPrincipal.menPorLet.Checked = True Then
            lstsort2.AddItem (EquivalenteLetra(VetorUsuario(cont, 1)) & " <= " & EquivalenteLetra(VetorUsuario(cont + 1, 1)))
        Else
            lstsort2.AddItem (EquivalenteCor(VetorUsuario(cont, 1)) & " <= " & EquivalenteCor(VetorUsuario(cont + 1, 1)))
        End If
    End If
Next cont
MsgBox "comaparacoes = " & comparacoes, vbInformation, "Ok !"
lblsort2.Caption = comparacoes & " comparações"
End Sub

Sub binn2()
Dim cont As Byte
Dim max As Byte
Dim min As Byte
Dim media As Byte
Dim comparacoes As Byte
Dim recont As Integer

For cont = 0 To 13
    VetorUsuario(cont, 0) = Vetor(cont)
Next cont

For cont = 1 To 13
    If cont = 1 Then
        comparacoes = comparacoes + 1
        If VetorUsuario(0, 0) > VetorUsuario(1, 0) Then
            
            Dim memoria As Byte
            memoria = VetorUsuario(0, 0)
            VetorUsuario(0, 0) = VetorUsuario(1, 0)
            VetorUsuario(1, 0) = memoria
        End If
    Else
        max = cont
        min = 0
        media = cont \ 2
        While (max - min) > 1
            comparacoes = comparacoes + 1
            If VetorUsuario(media, 0) < VetorUsuario(cont, 0) Then
                min = media
            ElseIf VetorUsuario(media, 0) > VetorUsuario(cont, 0) Then
                max = media
            Else
                max = media
                min = media - 1
            End If
            media = (max + min) / 2
        Wend
        
        If min = 0 Then
            'cont = 2
            Dim memoria2 As Byte
            Dim posicaoinicial As Byte
            comparacoes = comparacoes + 1
            If VetorUsuario(cont, 0) > VetorUsuario(min, 0) Then
                posicaoinicial = 1
            Else
                posicaoinicial = 0
            End If
            'VetorUsuario(0, 0) = 109
            memoria2 = VetorUsuario(cont, 0)
            For recont = cont To posicaoinicial + 1 Step -1
                VetorUsuario(recont, 0) = VetorUsuario(recont - 1, 0)
            Next recont
            VetorUsuario(posicaoinicial, 0) = memoria2
        ElseIf max = cont Then
            Dim memoria3 As Byte
            Dim posicaofinal As Byte
            comparacoes = comparacoes + 1
            If VetorUsuario(cont, 0) < VetorUsuario(cont - 1, 0) Then
                memoria = VetorUsuario(cont, 0)
                VetorUsuario(cont, 0) = VetorUsuario(max, 0)
                VetorUsuario(max, 0) = memoria3
            End If
        Else
            Dim memoria4
            memoria4 = VetorUsuario(cont, 0)
            For recont = cont To max + 1 Step -1
                VetorUsuario(recont, 0) = VetorUsuario(recont - 1, 0)
            Next recont
            VetorUsuario(max, 0) = memoria4
        End If
    End If
Next cont
For cont = 0 To 13
    Debug.Print cont; " = "; VetorUsuario(cont, 0)
Next cont
Debug.Print "comparacoes = " & comparacoes
lbln2.Caption = comparacoes & " comparacoes"
End Sub

Sub shell()
Dim media As Byte
Dim X As Byte
Dim h As Integer
Dim t As Byte
Dim v As Byte
Dim vetorusuario1(1 To 14) As Byte
Dim comparacoes As Byte

For X = 0 To 13
    vetorusuario1(X + 1) = Vetor(X)
Next X

media = 7
While media <> 0
   For X = 1 To 14 - media
      h = X
      While h >= 1
         v = h + media
         comparacoes = comparacoes + 1
         If vetorusuario1(h) < vetorusuario1(v) Then
            'lstsort2.AddItem equivalentecor(vetorusuario1
            
            t = vetorusuario1(h)
            vetorusuario1(h) = vetorusuario1(v)
            vetorusuario1(v) = t
            h = h - media
         Else
            h = 0
         End If
      Wend
    Next X
    media = media \ 2
Wend

For X = 1 To 14
    Debug.Print X & " = " & vetorusuario1(X)
Next X
lblshell.Caption = comparacoes
End Sub

Sub quick()

s = 0
f = 1
l = 14

teste2 = True

For X = 0 To 13
    vetorusuario1(X + 1) = Vetor(X)
Next X

m = vetorusuario1((l + f) \ 2)
i = f
j = l

While teste2 = True
    While (vetorusuario1(i) < m)
        comparacoes = comparacoes + 1
        i = i + 1
    Wend
    While (vetorusuario1(j) > m)
        comparacoes = comparacoes + 1
        j = j - 1
    Wend
    If i <= j Then
        If i <> j Then
            aux = vetorusuario1(j)
            vetorusuario1(j) = vetorusuario1(i)
            vetorusuario1(i) = aux
        End If
        i = i + 1
        j = j - 1
        If i <= j Then
            i = i + 1
        Else
            final
        End If
    Else
        final
    End If
Wend
For X = 1 To 14
    lstquick.AddItem (X & " = " & vetorusuario1(X))
Next X
lblquick.Caption = comparacoes & " comparacoes"
End Sub

Sub final()
    If i < l Then
        f1(s) = i
        l1(s) = l
        s = s + 1
    End If
    l = j
    If f >= l Then
        If s = 0 Then
            teste2 = False
        Else
            s = s - 1
            f = f1(s)
            l = l1(s)
        End If
    End If
m = vetorusuario1((l + f) \ 2)
i = f
j = l
End Sub

