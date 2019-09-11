VERSION 5.00
Begin VB.Form formComparacaoAlgoritmos 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenação - comparação de algoritmos"
   ClientHeight    =   5550
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8685
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVoltar 
      BackColor       =   &H0000C000&
      Caption         =   "Voltar"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H0000C000&
      Caption         =   "Ok"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtComentario 
      Height          =   1695
      Left            =   1440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   960
      Width           =   6855
   End
   Begin VB.ListBox lstProcesso 
      Height          =   1815
      Left            =   1440
      TabIndex        =   5
      Top             =   3480
      Width           =   6855
   End
   Begin VB.ComboBox cmbAlgoritmo 
      Height          =   315
      ItemData        =   "Form1.frx":030A
      Left            =   2040
      List            =   "Form1.frx":031D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lbltotal 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Selecione o algoritmo:"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Processo:"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comentário:"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.Menu menOpcoes 
      Caption         =   "&Opções"
      Begin VB.Menu menSort 
         Caption         =   "Sort"
      End
      Begin VB.Menu menSortII 
         Caption         =   "Sort II"
      End
      Begin VB.Menu menBolha 
         Caption         =   "Bolha"
      End
      Begin VB.Menu menShell 
         Caption         =   "Shell sort"
      End
      Begin VB.Menu menOrdv2 
         Caption         =   "OrdV2"
      End
      Begin VB.Menu menBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu menVoltar 
         Caption         =   "Voltar"
      End
   End
   Begin VB.Menu menAjuda 
      Caption         =   "&Ajuda"
      Begin VB.Menu menTopicos 
         Caption         =   "Tópicos de Ajuda"
         Shortcut        =   {F1}
      End
      Begin VB.Menu menResumo 
         Caption         =   "Resumo desta tela"
      End
      Begin VB.Menu menSobre 
         Caption         =   "Sobre..."
      End
   End
End
Attribute VB_Name = "formComparacaoAlgoritmos"
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


Private Sub cmdVoltar_Click()
   Unload Me
End Sub

'Private Sub Form_DblClick()
'   sort3
'End Sub

Private Sub menBolha_Click()
   cmbAlgoritmo.ListIndex = 2
   bolha
End Sub

Private Sub menOrdv2_Click()
   cmbAlgoritmo.ListIndex = 4
   OrdV2
End Sub

Private Sub menShell_Click()
   cmbAlgoritmo.ListIndex = 3
   shell
End Sub

Private Sub menSobre_Click()
   formSobre.Show 1
End Sub

Private Sub mensort_Click()
    cmbAlgoritmo.ListIndex = 0 ' (cmbAlgoritmo.ListIndex) = "Sort"
    sort
End Sub

Private Sub menSortII_Click()
    cmbAlgoritmo.ListIndex = 1
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
            lstProcesso.AddItem (EquivalenteLetra(VetorUsuario(cont, 1)) & " > " & EquivalenteLetra(VetorUsuario(cont + 1, 1))) & " - Trocou e reiniciou"
        Else
            lstProcesso.AddItem (EquivalenteCor(VetorUsuario(cont, 1)) & " > " & EquivalenteCor(VetorUsuario(cont + 1, 1))) & " - Trocou e reiniciou"
        End If
        aux = VetorUsuario(cont + 1, 1)
        VetorUsuario(cont + 1, 1) = VetorUsuario(cont, 1)
        VetorUsuario(cont, 1) = aux
        cont = -1
    Else
        If FormPrincipal.menPorLet.Checked = True Then
            lstProcesso.AddItem (EquivalenteLetra(VetorUsuario(cont, 1)) & " <= " & EquivalenteLetra(VetorUsuario(cont + 1, 1)))
        Else
            lstProcesso.AddItem (EquivalenteCor(VetorUsuario(cont, 1)) & " <= " & EquivalenteCor(VetorUsuario(cont + 1, 1)))
        End If
    End If
Next cont
'MsgBox "comaparacoes = " & comparacoes, vbInformation, "Ok !"
lbltotal.Caption = comparacoes & " comparações"
txtComentario.Text = "O algoritmo 'Sort' compara o valor de uma célula com a seguinte, trocando as posições caso não estejam em ordem. A cada troca o processo é reiniciado, ocorrendo a repetição de comparações. Com o decorrer das trocas, os valores mais altos vão sendo fixados nas últimas células até a ordenação completa. É lento e rudimentar"

Dim texto As String
For cont = 0 To 13
    If FormPrincipal.menPorLet.Checked = True Then
        texto = texto & cont + 1 & "º = " & EquivalenteLetra(VetorUsuario(cont, 1)) & "; valor = " & VetorUsuario(cont, 0) & vbCrLf
    Else
        texto = texto & cont + 1 & "º = " & EquivalenteCor(VetorUsuario(cont, 1)) & "; valor = " & VetorUsuario(cont, 0) & vbCrLf
    End If
Next cont

MsgBox "Foram necessárias " & comparacoes & " comparacoes. A ordem obtida foi:" & vbCrLf & vbCrLf & texto, vbInformation, "Processo concluído"
'Debug.Print "comparacoes = " & comparacoes


'Dim X As Byte
'For X = 0 To 13
'    Debug.Print X & " = " & VetorUsuario(X, 0)
'Next X
End Sub


Sub bolha()
Dim cont As Integer
Dim recont As Byte
Dim aux As Byte
Dim comparacoes As Integer
'Dim ok As Integer
Dim eq As Byte 'Para acertar o preenchimento do listbox
Dim v(1 To 14, 0 To 1) As Integer

For cont = 1 To 14
    v(cont, 0) = Vetor(cont - 1) ' Preenchendo o vetor usuario com os valores do vetor
    v(cont, 1) = cont - 1
Next cont
cont = 0

For cont = 1 To 13
    eq = cont
    For recont = 1 To (14 - cont)
        comparacoes = comparacoes + 1
        If Not (v(recont, 0) < v(recont + 1, 0)) Then
            'ok = ok + 1
            aux = v(recont, 0)
            v(recont, 0) = v(recont + 1, 0)
            v(recont + 1, 0) = aux
            
            aux = v(recont, 1)
            v(recont, 1) = v(recont + 1, 1)
            v(recont + 1, 1) = aux
            If FormPrincipal.menPorLet.Checked = True Then
                lstProcesso.AddItem (EquivalenteLetra(v(recont + 1, 1)) & " > " & EquivalenteLetra(v(recont, 1))) & " - Trocou os 2"
            Else
                lstProcesso.AddItem (EquivalenteCor(v(recont + 1, 1)) & " > " & EquivalenteCor(v(recont, 1))) & " - Trocou os 2"
            End If
            
         eq = recont
         Else
            If FormPrincipal.menPorLet.Checked = True Then
                lstProcesso.AddItem (EquivalenteLetra(v(recont, 1)) & " <= " & EquivalenteLetra(v(recont + 1, 1)))
            Else
                lstProcesso.AddItem (EquivalenteCor(v(recont, 1)) & " <= " & EquivalenteCor(v(recont + 1, 1)))
            End If
        End If
    Next 'recont
lstProcesso.List(lstProcesso.ListCount - 1) = lstProcesso.List(lstProcesso.ListCount - 1) & " - Achou o " & 14 - cont + 1 & "º"
Next 'cont
'MsgBox "comaparacoes = " & comparacoes, vbInformation, "Ok !"
lbltotal.Caption = comparacoes & " comparações"
txtComentario.Text = "Método clássico que inicia comparando as duas primeiras células e coloca-as em ordem. Em seguida compara a maior destas com a 3ª e assim por diante, fixando o valor mais alto na última célula. Repete o processo até a penúltima, depois até a antipenúltima. Termina colocando os 2 menores valores nas primeiras posições. Sempre executa o mesmo número de comparações para uma mesma quantidade de dados a ser ordenada"

Dim texto As String
For cont = 1 To 14
    If FormPrincipal.menPorLet.Checked = True Then
        texto = texto & cont & "º = " & EquivalenteLetra(v(cont, 1)) & "; valor = " & v(cont, 0) & vbCrLf
    Else
        texto = texto & cont & "º = " & EquivalenteCor(v(cont, 1)) & "; valor = " & v(cont, 0) & vbCrLf
    End If
Next cont



MsgBox "Foram necessárias " & comparacoes & " comparacoes. A ordem obtida foi:" & vbCrLf & vbCrLf & texto, vbInformation, "Processo concluído"
'Debug.Print "comparacoes = " & comparacoes


'Dim X As Byte
'For X = 1 To 14
'    Debug.Print X & " = " & v(X, 0)
'Next X

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
            lstProcesso.AddItem (EquivalenteLetra(VetorUsuario(cont + 1, 1)) & " > " & EquivalenteLetra(VetorUsuario(cont, 1))) & " - Troca os 2"
        Else
            lstProcesso.AddItem (EquivalenteCor(VetorUsuario(cont + 1, 1)) & " > " & EquivalenteCor(VetorUsuario(cont, 1))) & " - Troca os 2"
        End If
        
        If cont <> 0 Then
            lstProcesso.List(lstProcesso.ListCount - 1) = lstProcesso.List(lstProcesso.ListCount - 1) & " e compara com os anteriores"
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
                        lstProcesso.AddItem (EquivalenteLetra(VetorUsuario(recont, 1)) & " > " & EquivalenteLetra(VetorUsuario(recont - 1, 1))) & " - Troca os 2"
                    Else
                        lstProcesso.AddItem (EquivalenteCor(VetorUsuario(recont, 1)) & " > " & EquivalenteCor(VetorUsuario(recont - 1, 1))) & " - Troca os 2"
                    End If
                 Else
                    If FormPrincipal.menPorLet.Checked = True Then
                        lstProcesso.AddItem (EquivalenteLetra(VetorUsuario(recont - 1, 1)) & " <= " & EquivalenteLetra(VetorUsuario(recont, 1)))
                    Else
                        lstProcesso.AddItem (EquivalenteCor(VetorUsuario(recont - 1, 1)) & " <= " & EquivalenteCor(VetorUsuario(recont, 1)))
                    End If
                    recont = 1
                End If
            Next recont
        End If
    Else
        If FormPrincipal.menPorLet.Checked = True Then
            lstProcesso.AddItem (EquivalenteLetra(VetorUsuario(cont, 1)) & " <= " & EquivalenteLetra(VetorUsuario(cont + 1, 1)))
        Else
            lstProcesso.AddItem (EquivalenteCor(VetorUsuario(cont, 1)) & " <= " & EquivalenteCor(VetorUsuario(cont + 1, 1)))
        End If
    End If
Next cont
'MsgBox "comaparacoes = " & comparacoes, vbInformation, "Ok !"
lbltotal.Caption = comparacoes & " comparações"
txtComentario.Text = "Atua 'jogando' os valores menores para as primeiras células, começando a comparação pelas primeiras células e indo uma a uma até a última. Muito mais eficiente que o 'Sort'"

Dim texto As String
For cont = 0 To 13
    If FormPrincipal.menPorLet.Checked = True Then
        texto = texto & cont + 1 & "º = " & EquivalenteLetra(VetorUsuario(cont, 1)) & "; valor = " & VetorUsuario(cont, 0) & vbCrLf
    Else
        texto = texto & cont + 1 & "º = " & EquivalenteCor(VetorUsuario(cont, 1)) & "; valor = " & VetorUsuario(cont, 0) & vbCrLf
    End If
Next cont
MsgBox "Foram necessárias " & comparacoes & " comparacoes. A ordem obtida foi:" & vbCrLf & vbCrLf & texto, vbInformation, "Processo concluído"
'Debug.Print "comparacoes = " & comparacoes



'Dim X As Byte
'For X = 0 To 13
'    Debug.Print X & " = " & VetorUsuario(X, 0)
'Next X

End Sub

Sub OrdV2()
Dim k As Byte
Dim cont As Byte
Dim max As Byte
Dim min As Byte
Dim media As Byte
Dim comparacoes As Byte
Dim recont As Integer

For cont = 0 To 13
    VetorUsuario(cont, 0) = Vetor(cont)
    VetorUsuario(cont, 1) = cont
Next cont

For cont = 1 To 13
    If cont = 1 Then
        comparacoes = comparacoes + 1
        If VetorUsuario(0, 0) > VetorUsuario(1, 0) Then
            Dim memoria As Byte
            For k = 0 To 1
               memoria = VetorUsuario(0, k)
               VetorUsuario(0, k) = VetorUsuario(1, k)
               VetorUsuario(1, k) = memoria
            Next
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
            Dim memoria2(0 To 1) As Byte
            Dim posicaoinicial As Byte
            comparacoes = comparacoes + 1
            If VetorUsuario(cont, 0) > VetorUsuario(min, 0) Then
                posicaoinicial = 1
            Else
                posicaoinicial = 0
            End If
            'VetorUsuario(0, 0) = 109
            memoria2(0) = VetorUsuario(cont, 0)
            memoria2(1) = VetorUsuario(cont, 1)
            For recont = cont To posicaoinicial + 1 Step -1
                VetorUsuario(recont, 0) = VetorUsuario(recont - 1, 0)
                VetorUsuario(recont, 1) = VetorUsuario(recont - 1, 1)
            Next recont
            VetorUsuario(posicaoinicial, 0) = memoria2(0)
            VetorUsuario(posicaoinicial, 1) = memoria2(1)
        ElseIf max = cont Then
            Dim memoria3 As Byte
            Dim posicaofinal As Byte
            comparacoes = comparacoes + 1
            If VetorUsuario(cont, 0) < VetorUsuario(cont - 1, 0) Then
               For k = 0 To 1
                  memoria = VetorUsuario(cont, k)
                  VetorUsuario(cont, k) = VetorUsuario(max, k)
                  VetorUsuario(max, k) = memoria3
               Next
            End If
        Else
            Dim memoria4(0 To 1) As Byte
            memoria4(0) = VetorUsuario(cont, 0)
            memoria4(1) = VetorUsuario(cont, 1)
            For recont = cont To max + 1 Step -1
                VetorUsuario(recont, 0) = VetorUsuario(recont - 1, 0)
                VetorUsuario(recont, 1) = VetorUsuario(recont - 1, 1)
            Next recont
            VetorUsuario(max, 0) = memoria4(0)
            VetorUsuario(max, 1) = memoria4(1)
        End If
    End If
Next cont

lbltotal.Caption = comparacoes & " comparacoes"
txtComentario.Text = "Algoritmo desenvolvido com este software. Começa comparando os três primeiros valores e os ordena, fazendo uma lista conhecida. A partir daí compara os novos valores sempre com o meio da lista conhecida, inicialmente descobrindo em qual metade (inferior ou superior) o registro deva ser colocado. Tendo registros conhecidos suficientes, vai 'quebrando' a parte da lista descoberta até encontrar a posição correta para o registo desconhecido. Ex: descobre em qual metade colocar o registro, em seguida, em qual metade desta metada (ou seja, em qual 'quarto' da lista conhecida) e assim por diante."
lstProcesso.AddItem ("Detalhes do processo na próxima versão")

Dim texto As String
For cont = 0 To 13
    If FormPrincipal.menPorLet.Checked = True Then
        texto = texto & cont + 1 & "º = " & EquivalenteLetra(VetorUsuario(cont, 1)) & "; valor = " & VetorUsuario(cont, 0) & vbCrLf
    Else
        texto = texto & cont + 1 & "º = " & EquivalenteCor(VetorUsuario(cont, 1)) & "; valor = " & VetorUsuario(cont, 0) & vbCrLf
    End If
Next cont

MsgBox "Foram necessárias " & comparacoes & " comparacoes. A ordem obtida foi:" & vbCrLf & vbCrLf & texto, vbInformation, "Processo concluído"
'Debug.Print "comparacoes = " & comparacoes

End Sub

Sub shell()
Dim media As Byte
Dim X As Byte
Dim h As Integer
Dim t As Byte
Dim v As Byte
Dim vetorShell(1 To 14, 0 To 1) As Byte
Dim comparacoes As Byte

For X = 0 To 13
    vetorShell(X + 1, 0) = Vetor(X)
    vetorShell(X + 1, 1) = X
Next X

media = 7
'media = 6
While media <> 0
   For X = 1 To 14 - media
      h = X
      While h >= 1
         v = h + media
         comparacoes = comparacoes + 1
         If vetorShell(h, 0) > vetorShell(v, 0) Then
            'lstsort2.AddItem equivalentecor(vetorusuario1
             If FormPrincipal.menPorLet.Checked = True Then
                lstProcesso.AddItem (EquivalenteLetra(vetorShell(h, 1)) & " > " & EquivalenteLetra(vetorShell(v, 1)) & " - troca os 2")
             Else
                lstProcesso.AddItem (EquivalenteCor(vetorShell(h, 1)) & " > " & EquivalenteCor(vetorShell(v, 1)) & " - troca os 2")
             End If
            
            t = vetorShell(h, 0)
            vetorShell(h, 0) = vetorShell(v, 0)
            vetorShell(v, 0) = t
            
            t = vetorShell(h, 1)
            vetorShell(h, 1) = vetorShell(v, 1)
            vetorShell(v, 1) = t
            h = h - media
         Else
             If FormPrincipal.menPorLet.Checked = True Then
                lstProcesso.AddItem (EquivalenteLetra(vetorShell(h, 1)) & " < " & EquivalenteLetra(vetorShell(v, 1)))
             Else
                lstProcesso.AddItem (EquivalenteCor(vetorShell(h, 1)) & " < " & EquivalenteCor(vetorShell(v, 1)))
             End If
            h = 0
         End If
      Wend
    Next X
    media = media \ 2
Wend

'For X = 1 To 14
'    Debug.Print X & " = " & vetorShell(X, 0)
'Next X
lbltotal.Caption = comparacoes & " comparações"
txtComentario.Text = "Evita repetições, utilizando muitas comparações indiretas (ex: se A>B e B>C, certamente A>C, não é necessária uma comparação para conluir isto). Uma análise com o fluxograma (do arquivo de ajuda) e a lista abaixo leva ao entendimento do processo em detalhes"

'lbltotal.Caption = comparacoes & " comparacoes"
Dim texto As String
Dim cont As Byte
For cont = 1 To 14
    If FormPrincipal.menPorLet.Checked = True Then
        texto = texto & cont & "º = " & EquivalenteLetra(vetorShell(cont, 1)) & "; valor = " & vetorShell(cont, 0) & vbCrLf
    Else
        texto = texto & cont & "º = " & EquivalenteCor(vetorShell(cont, 1)) & "; valor = " & vetorShell(cont, 0) & vbCrLf
    End If
Next cont
MsgBox "Foram necessárias " & comparacoes & " comparacoes. A ordem obtida foi:" & vbCrLf & vbCrLf & texto, vbInformation, "Processo concluído"
'Debug.Print "comparacoes = " & comparacoes
End Sub

Private Sub cmdOk_Click()
If cmbAlgoritmo.List(cmbAlgoritmo.ListIndex) = "" Then
   MsgBox "Favor selecione um algoritmo de ordenação", vbExclamation, "Erro"
Else
   lstProcesso.Clear
   If cmbAlgoritmo.List(cmbAlgoritmo.ListIndex) = "Sort" Then
      sort
   ElseIf cmbAlgoritmo.List(cmbAlgoritmo.ListIndex) = "Sort tipo II" Then
      sort2
   ElseIf cmbAlgoritmo.List(cmbAlgoritmo.ListIndex) = "Bolha" Then
      bolha
   ElseIf cmbAlgoritmo.List(cmbAlgoritmo.ListIndex) = "Shell Sort" Then
      shell
   ElseIf cmbAlgoritmo.List(cmbAlgoritmo.ListIndex) = "OrdV2" Then
      OrdV2
   End If
End If
End Sub

Private Sub menResumo_Click()
   MsgBox "Nesta tela você pode comparar seu desempenho com o desempenho de algoritmos simples de ordenação, pois os algoritmos ordenarão os mesmos dados apresentados na tela principal. Basta selecionar o algoritmo e verificar seu desempenho. Mais informações, como por exemplo fluxogramas dos algoritmos, na seção 'Ajuda'.", vbInformation, "Ajuda resumida"
End Sub

Private Sub menTopicos_Click()
    Dim strArquivo As String
    Dim objAjuda As ClasseAjuda
 
    Set objAjuda = New ClasseAjuda
    
    strArquivo = App.Path & "\ajuda\ajuda.chm"
    Call objAjuda.Show(strArquivo) ', "janelaHelp")
    Set objAjuda = Nothing
End Sub

Private Sub menVoltar_Click()
   Unload Me
End Sub





























'Sub sort3()
'Dim cont As Integer
'Dim recont As Integer
'Dim aux As Byte
'Dim maxcont As Byte
'Dim comparacoes As Integer
''Dim ok As Integer
'Dim v(1 To 14, 0 To 1) As Byte'
'
'For cont = 0 To 13
'    v(cont + 1, 0) = Vetor(cont) ' Preenchendo o vetor usuario com os valores do vetor
'    v(cont + 1, 1) = cont
'Next cont
'cont = 0'
'''
'
'For cont = 1 To 13
'    comparacoes = comparacoes + 1
'    If v(cont, 0) > v(cont + 1, 0) Then
' '       ok = ok + 1
'        aux = v(cont, 0)
'        v(cont, 0) = v(cont + 1, 0)
'        v(cont + 1, 0) = aux
'
'        aux = VetorUsuario(cont, 1)
'        v(cont, 1) = v(cont + 1, 1)
'        v(cont + 1, 1) = aux
'
'        If FormPrincipal.menPorLet.Checked = True Then
'            lstProcesso.AddItem (EquivalenteLetra(v(cont + 1, 1)) & " > " & EquivalenteLetra(v(cont, 1))) & " - Troca os 2"
'        Else
'            lstProcesso.AddItem (EquivalenteCor(v(cont + 1, 1)) & " > " & EquivalenteCor(v(cont, 1))) & " - Troca os 2"
'        End If
'
'        If cont <> 1 Then
'            lstProcesso.List(lstProcesso.ListCount - 1) = lstProcesso.List(lstProcesso.ListCount - 1) & " e compara com os anteriores"
'            For recont = cont To 2 Step -1
'                comparacoes = comparacoes + 1
'                If v(recont - 1, 0) > v(recont, 0) Then
'  '                  ok = ok + 1
'                    aux = v(recont, 0)
'                    v(recont, 0) = v(recont - 1, 0)
'                    v(recont - 1, 0) = aux
'
'                    aux = v(recont, 1)
'                    v(recont, 1) = v(recont - 1, 1)
'                    v(recont - 1, 1) = aux
'
'                    If FormPrincipal.menPorLet.Checked = True Then
'                        lstProcesso.AddItem (EquivalenteLetra(v(recont, 1)) & " > " & EquivalenteLetra(v(recont - 1, 1))) & " - Troca os 2"
'                    Else
'                        lstProcesso.AddItem (EquivalenteCor(v(recont, 1)) & " > " & EquivalenteCor(v(recont - 1, 1))) & " - Troca os 2"
'                    End If
'                 Else
'                    If FormPrincipal.menPorLet.Checked = True Then
'                        lstProcesso.AddItem (EquivalenteLetra(v(recont - 1, 1)) & " <= " & EquivalenteLetra(v(recont, 1)))
'                    Else
'                        lstProcesso.AddItem (EquivalenteCor(v(recont - 1, 1)) & " <= " & EquivalenteCor(v(recont, 1)))
'                    End If
'                    recont = 2
'                End If
'            Next recont
'        End If
'    Else
'        If FormPrincipal.menPorLet.Checked = True Then
'            lstProcesso.AddItem (EquivalenteLetra(v(cont, 1)) & " <= " & EquivalenteLetra(v(cont + 1, 1)))
'        Else
'            lstProcesso.AddItem (EquivalenteCor(v(cont, 1)) & " <= " & EquivalenteCor(v(cont + 1, 1)))
'        End If
'    End If
'Next cont
''MsgBox "comaparacoes = " & comparacoes, vbInformation, "Ok !"
'lbltotal.Caption = comparacoes & " comparações"
'txtComentario.Text = "Atua 'jogando' os valores menores para as primeiras células, começando a comparação pelas primeiras células e indo uma a uma até a última. Muito mais eficiente que o 'Sort'"'
'
'Dim texto As String
'For cont = 1 To 14
'    If FormPrincipal.menPorLet.Checked = True Then
'        texto = texto & cont + 1 & "º = " & EquivalenteLetra(v(cont, 1)) & "; valor = " & v(cont, 0) & vbCrLf
'    Else
'        texto = texto & cont + 1 & "º = " & EquivalenteCor(v(cont, 1)) & "; valor = " & v(cont, 0) & vbCrLf
'    End If
'Next cont
'MsgBox "Foram necessárias " & comparacoes & " comparacoes. A ordem obtida foi:" & vbCrLf & vbCrLf & texto, vbInformation, "Processo concluído"
''Debug.Print "comparacoes = " & comparacoes'



''Dim X As Byte
''For X = 0 To 13
''    Debug.Print X & " = " & VetorUsuario(X, 0)
''Next X'
'
'End Sub
