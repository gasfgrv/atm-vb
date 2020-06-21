Dim estatisticas(0 To 99, 0 To 2) As Integer

Dim notas(0 To 5) As Integer
Dim qtdNotas(0 To 5) As Integer

Dim TotalNotas As Integer
Dim valorTotal As Integer
Dim banco As Integer
Dim qtdSaque As Integer

Sub valorNotas()
    notas(0) = 100
    notas(1) = 50
    notas(2) = 20
    notas(3) = 10
    notas(4) = 5
    notas(5) = 2
End Sub

Private Sub CampoVazio()
    If (TextBox1.Text = "") Then
        TextBox1.Value = 0
    End If

    If (TextBox2.Text = "") Then
        TextBox2.Value = 0
    End If

    If (TextBox3.Text = "") Then
        TextBox3.Value = 0
    End If

    If (TextBox4.Text = "") Then
        TextBox4.Value = 0
    End If

    If (TextBox5.Text = "") Then
        TextBox5.Value = 0
    End If

    If (TextBox6.Text = "") Then
        TextBox6.Value = 0
    End If
End Sub

Public Sub CarregarNotas()
    valorNotas
    
    valorTotal = 0
    
    qtdNotas(0) = qtdNotas(0) + TextBox1.Value
    qtdNotas(1) = qtdNotas(1) + TextBox2.Value
    qtdNotas(2) = qtdNotas(2) + TextBox3.Value
    qtdNotas(3) = qtdNotas(3) + TextBox4.Value
    qtdNotas(4) = qtdNotas(4) + TextBox5.Value
    qtdNotas(5) = qtdNotas(5) + TextBox6.Value
    
    For i = 0 To 5
        valorTotal = valorTotal + (notas(i) * qtdNotas(i))
        TotalNotas = TotalNotas + qtdNotas(i)
    Next i
    
    MsgBox "Notas de R$" & notas(0) & ",00: " & qtdNotas(0) & vbNewLine & "Notas de R$" & notas(1) & ",00: " & qtdNotas(1) & vbNewLine & "Notas de R$" & notas(2) & ",00: " & qtdNotas(2) & vbNewLine & "Notas de R$" & notas(3) & ",00: " & qtdNotas(3) & vbNewLine & "Notas de R$" & notas(4) & ",00: " & qtdNotas(4) & vbNewLine & "Notas de R$" & notas(5) & ",00: " & qtdNotas(5) & vbNewLine & vbNewLine & "Valor Total: R$" & valorTotal & ",00", vbOKOnly, "Qtde de Notas"
End Sub

Private Function BancoSaque()
    If (OptionButton1.Value = True) Then
        BancoSaque = 1
    ElseIf (OptionButton2.Value = True) Then
        BancoSaque = 2
    ElseIf (OptionButton3.Value = True) Then
        BancoSaque = 4
    ElseIf (OptionButton4.Value = True) Then
        BancoSaque = 3
    End If
End Function

Public Sub RetirarNotas()
    Dim saque As Integer
    
    If (qtdSaque >= 100 Or TotalNotas < 1) Then
        MsgBox "Impossível sacar neste momento.", vbInformation, "Erro"
    Else
        'Escreva("1 - 1000,00 / 10 X 100");
        If (OptionButton5.Value = True And valorTotal >= 1000 And qtdNotas(0) >= 10) Then
            qtdNotas(0) = qtdNotas(0) - 10
            estatisticas(qtdSaque, 0) = qtdSaque + 1
            estatisticas(qtdSaque, 1) = BancoSaque
            estatisticas(qtdSaque, 2) = 1000
            qtdSaque = qtdSaque + 1
            MsgBox "Retire as notas.", vbInformation, "Saque Concluído"
        
        'Escreva("2 - 500,00 / 5 X 100");
        ElseIf (OptionButton6.Value = True And valorTotal >= 500 And qtdNotas(0) >= 5) Then
            qtdNotas(0) = qtdNotas(0) - 5
            estatisticas(qtdSaque, 0) = qtdSaque + 1
            estatisticas(qtdSaque, 1) = BancoSaque
            estatisticas(qtdSaque, 2) = 500
            qtdSaque = qtdSaque + 1
            MsgBox "Retire as notas.", vbInformation, "Saque Concluído"
        
        'Escreva("3 - 500,00 / 3 X 100, 2 X 50, 5 X 20");
        ElseIf (OptionButton7.Value = True And valorTotal >= 500 And qtdNotas(0) >= 3 And qtdNotas(1) >= 2 And qtdNotas(3) >= 5) Then
            qtdNotas(0) = qtdNotas(0) - 3
            qtdNotas(1) = qtdNotas(1) - 2
            qtdNotas(2) = qtdNotas(2) - 5
            estatisticas(qtdSaque, 0) = qtdSaque + 1
            estatisticas(qtdSaque, 1) = BancoSaque
            estatisticas(qtdSaque, 2) = 500
            qtdSaque = qtdSaque + 1
            MsgBox "Retire as notas.", vbInformation, "Saque Concluído"
        
        'Escreva("4 - 250,00 / 5 X 50");
        ElseIf (OptionButton8.Value = True And valorTotal >= 250 And qtdNotas(1) >= 5) Then
            qtdNotas(1) = qtdNotas(1) - 5
            estatisticas(qtdSaque, 0) = qtdSaque + 1
            estatisticas(qtdSaque, 1) = BancoSaque
            estatisticas(qtdSaque, 2) = 250
            qtdSaque = qtdSaque + 1
            MsgBox "Retire as notas.", vbInformation, "Saque Concluído"
        
        'Escreva("5 - 250,00 / 2 X 100, 5 X 10");
        ElseIf (OptionButton9.Value = True And valorTotal >= 250 And qtdNotas(0) >= 2 And qtdNotas(3) >= 5) Then
            qtdNotas(0) = qtdNotas(0) - 2
            qtdNotas(3) = qtdNotas(3) - 5
            estatisticas(qtdSaque, 0) = qtdSaque + 1
            estatisticas(qtdSaque, 1) = BancoSaque
            estatisticas(qtdSaque, 2) = 250
            qtdSaque = qtdSaque + 1
            MsgBox "Retire as notas.", vbInformation, "Saque Concluído"
        
        'Escreva("6 - 100,00 / 2 X 50");
        ElseIf (OptionButton10.Value = True And valorTotal >= 100 And qtdNotas(1) >= 2) Then
            qtdNotas(1) = qtdNotas(1) - 2
            estatisticas(qtdSaque, 0) = qtdSaque + 1
            estatisticas(qtdSaque, 1) = BancoSaque
            estatisticas(qtdSaque, 2) = 100
            qtdSaque = qtdSaque + 1
            MsgBox "Retire as notas.", vbInformation, "Saque Concluído"
        
        'Escreva("7 - 100,00 / 5 X 20");
        ElseIf (OptionButton11.Value = True And valorTotal >= 100 And qtdNotas(2) >= 5) Then
            qtdNotas(2) = qtdNotas(2) - 5
            estatisticas(qtdSaque, 0) = qtdSaque + 1
            estatisticas(qtdSaque, 1) = BancoSaque
            estatisticas(qtdSaque, 2) = 100
            qtdSaque = qtdSaque + 1
            MsgBox "Retire as notas.", vbInformation, "Saque Concluído"
        
        'Escreva("8 - 100,00 / 10 X 10");
        ElseIf (OptionButton12.Value = True And valorTotal >= 100 And qtdNotas(3) >= 10) Then
            qtdNotas(3) = qtdNotas(3) - 10
            estatisticas(qtdSaque, 0) = qtdSaque + 1
            estatisticas(qtdSaque, 1) = BancoSaque
            estatisticas(qtdSaque, 2) = 100
            qtdSaque = qtdSaque + 1
            MsgBox "Retire as notas.", vbInformation, "Erro"
        ElseIf (OptionButton13.Value = True) Then
            saque = TextBox7.Value
            
            If (saque > valorTotal Or saque < notas(5) Or saque = 3) Then
                MsgBox "Valor inválido.", vbInformation, "Erro"
            Else
                ContarNotas (saque)
                estatisticas(qtdSaque, 0) = qtdSaque + 1
                estatisticas(qtdSaque, 1) = BancoSaque
                estatisticas(qtdSaque, 2) = saque
                qtdSaque = qtdSaque + 1
            End If
        
        Else
            MsgBox "Opção inválida neste momento, por favor recarregue as notas", vbInformation, "Erro"
        End If
    End If
End Sub

Private Sub ContarNotas(s As Integer)
    Dim nNotas As Integer
    Dim msg As String
    
    valorNotas
    
    msg = ""
    
    For i = 0 To 5
        nNotas = s \ notas(i)
        
        If (nNotas <= qtdNotas(i)) Then
            If (nNotas > 0) Then
                qtdNotas(i) = qtdNotas(i) - nNotas
                msg = msg & nNotas & " Nota(s) de " & notas(i) & ",00" & vbCrLf
            End If
            
            s = s Mod notas(i)
        ElseIf qtdNotas(i) > 0 Then
            nNotas = (s - (notas(i) * qtdNotas(i)))
            nNotas = nNotas \ notas(i)
            
            If (nNotas > 0) Then
                qtdNotas(i) = qtdNotas(i) - nNotas
                msg = msg & nNotas & " Nota(s) de " & notas(i) & ",00" & vbCrLf
            End If
            
            s = s - (nNotas * notas(i))
        Else
            MsgBox "Não há notas Suficientes para o saque.", vbInformation
        End If
    
    Next i
    
    MsgBox msg, vbOKOnly, "Retire as Notas"
End Sub

Private Sub limparCampos()
    TextBox1.Text = ""
    TextBox2.Text = ""
    TextBox3.Text = ""
    TextBox4.Text = ""
    TextBox5.Text = ""
    TextBox6.Text = ""
    TextBox1.SetFocus
End Sub

Private Sub MaiorMenorvalor()
    Label11.Visible = True
    Label11.Caption = MenorMaiorSaque(1)
    
    Label12.Visible = True
    Label12.Caption = MenorMaiorSaque(2)
    
    Label13.Visible = True
    Label13.Caption = MenorMaiorSaque(3)
    
    Label14.Visible = True
    Label14.Caption = MenorMaiorSaque(4)
End Sub

Private Sub MediaSaques()
    Label11.Visible = True
    Label11.Caption = media(1)
    
    Label12.Visible = True
    Label12.Caption = media(2)
    
    Label13.Visible = True
    Label13.Caption = media(3)
    
    Label14.Visible = True
    Label14.Caption = media(4)
End Sub

Private Sub SomaSaques()
    Label11.Visible = True
    Label11.Caption = soma(1)
    
    Label12.Visible = True
    Label12.Caption = soma(2)
    
    Label13.Visible = True
    Label13.Caption = soma(3)
    
    Label14.Visible = True
    Label14.Caption = soma(4)
End Sub

Private Function MenorMaiorSaque(b As Integer)
    Dim maior As Integer
    Dim menos As Integer
    Dim contSaque As Integer

    maior = 0
    menor = valorTotal
    contSaque = 0

    For i = 0 To 99
        For j = 0 To 2
            If (estatisticas(i, 1) = b) Then
                contSaque = contSaque + 1
                If (estatisticas(i, 2) > maior) Then
                    maior = estatisticas(i, 2)
                End If

                If (estatisticas(i, 2) < menor) Then
                    menor = estatisticas(i, 2)
                End If
            End If
        Next j
    Next i
    
    If contSaque < 1 Then
        menor = 0
    End If
    
    If contSaque = 1 Then
        menor = maior
    End If
    
    MenorMaiorSaque = "Maior saque: R$" & maior & ",00." & vbCrLf & "Menor saque: R$" & menor & ",00."
End Function

Private Function media(b As Integer)
    Dim soma As Integer
    Dim contSaque As Integer
    
    soma = 0
    contSaque = 0
    
    For i = 0 To 99
        If (estatisticas(i, 1) = b) Then
            soma = soma + estatisticas(i, 2)
            contSaque = contSaque + 1
        End If
    Next i
    
    If contSaque < 1 Then
        media = "Soma dos valores do saque: R$" & 0 & ",00"
    Else
        media = "Soma dos valores do saque: R$" & soma \ contSaque & ",00"
    End If
End Function

Private Function soma(b As Integer)
    Dim contSaque As Integer
    
    soma = 0
    contSaque = 0
    
    For i = 0 To 99
        If (estatisticas(i, 1) = b) Then
            soma = soma + estatisticas(i, 2)
            contSaque = contSaque + 1
        End If
    Next i
    
    If contSaque < 1 Then
        soma = "Soma dos valores do saque: R$" & 0 & ",00"
    Else
        soma = "Soma dos valores do saque: R$" & soma & ",00"
    End If
End Function

Private Sub atualizarEstatisticas()
    If OptionButton14.Value = True Then
        MaiorMenorvalor
    ElseIf OptionButton15.Value = True Then
        MediaSaques
    ElseIf OptionButton16.Value = True Then
        SomaSaques
    ElseIf OptionButton14.Value = False And OptionButton15.Value = False And OptionButton16.Value = False Then
        MsgBox "Selecione uma da opções", vbInformation, Erro
    End If
End Sub

Private Sub CommandButton1_Click()
    limparCampos
End Sub

Private Sub CommandButton2_Click()
    CampoVazio
    CarregarNotas
    limparCampos
End Sub

Private Sub CommandButton3_Click()
    OptionButton1.Value = True
    OptionButton5.Value = True
End Sub

Private Sub CommandButton4_Click()
    RetirarNotas
End Sub

Private Sub CommandButton5_Click()
    Dim sobras As Integer
    
    sobras = 0
    
    For i = 0 To 5
        sobras = sobras + (notas(i) * qtdNotas(i))
    Next i
    
    MsgBox "Notas de R$" & notas(0) & ",00: " & qtdNotas(0) & vbNewLine & "Notas de R$" & notas(1) & ",00: " & qtdNotas(1) & vbNewLine & "Notas de R$" & notas(2) & ",00: " & qtdNotas(2) & vbNewLine & "Notas de R$" & notas(3) & ",00: " & qtdNotas(3) & vbNewLine & "Notas de R$" & notas(4) & ",00: " & qtdNotas(4) & vbNewLine & "Notas de R$" & notas(5) & ",00: " & qtdNotas(5) & vbNewLine & vbNewLine & "Valor Total: R$" & sobras & ",00", vbOKOnly, "Qtde de Notas"
End Sub

Private Sub CommandButton6_Click()
    If MsgBox("Deseja Sair", vbYesNo, "Caixa Eletrônico") = vbYes Then
        Unload UserForm1
        ThisWorkbook.Save
    End If
End Sub

Private Sub CommandButton7_Click()
    atualizarEstatisticas
End Sub

Private Sub OptionButton10_Click()
    TextBox7.Text = "0"
    TextBox7.Enabled = False
End Sub

Private Sub OptionButton11_Click()
    TextBox7.Text = "0"
    TextBox7.Enabled = False
End Sub

Private Sub OptionButton12_Click()
    TextBox7.Text = "0"
    TextBox7.Enabled = False
End Sub

Private Sub OptionButton13_Click()
    TextBox7.Enabled = True
End Sub

Private Sub OptionButton14_Click()
    MaiorMenorvalor
End Sub

Private Sub OptionButton15_Click()
    MediaSaques
End Sub

Private Sub OptionButton16_Click()
    SomaSaques
End Sub

Private Sub OptionButton5_Click()
    TextBox7.Text = "0"
    TextBox7.Enabled = False
End Sub

Private Sub OptionButton6_Click()
    TextBox7.Text = "0"
    TextBox7.Enabled = False
End Sub

Private Sub OptionButton7_Click()
    TextBox7.Text = "0"
    TextBox7.Enabled = False
End Sub

Private Sub OptionButton8_Click()
    TextBox7.Text = "0"
    TextBox7.Enabled = False
End Sub

Private Sub OptionButton9_Click()
    TextBox7.Text = "0"
    TextBox7.Enabled = False
End Sub

Private Sub TextBox1_Change()
    If Len(TextBox1) > 0 Then
        tamanho = Len(TextBox1)
        ultimo = Right(TextBox1, 1)
        
        If Not IsNumeric(ultimo) Then
            TextBox1 = Left(TextBox1, tamanho - 1)
        End If
    End If
End Sub

Private Sub TextBox2_Change()
    If Len(TextBox2) > 0 Then
        tamanho = Len(TextBox2)
        ultimo = Right(TextBox2, 1)
        
        If Not IsNumeric(ultimo) Then
            TextBox2 = Left(TextBox2, tamanho - 1)
        End If
    End If
End Sub

Private Sub TextBox3_Change()
    If Len(TextBox3) > 0 Then
        tamanho = Len(TextBox3)
        ultimo = Right(TextBox3, 1)
       
        If Not IsNumeric(ultimo) Then
            TextBox3 = Left(TextBox3, tamanho - 1)
        End If
    End If
End Sub

Private Sub TextBox4_Change()
    If Len(TextBox4) > 0 Then
        tamanho = Len(TextBox4)
        ultimo = Right(TextBox4, 1)
      
        If Not IsNumeric(ultimo) Then
            TextBox4 = Left(TextBox4, tamanho - 1)
        End If
    End If
End Sub

Private Sub TextBox5_Change()
    If Len(TextBox5) > 0 Then
        tamanho = Len(TextBox5)
        ultimo = Right(TextBox5, 1)
      
        If Not IsNumeric(ultimo) Then
            TextBox5 = Left(TextBox5, tamanho - 1)
        End If
    End If
End Sub

Private Sub TextBox6_Change()
    If Len(TextBox6) > 0 Then
        tamanho = Len(TextBox6)
        ultimo = Right(TextBox6, 1)
     
        If Not IsNumeric(ultimo) Then
            TextBox6 = Left(TextBox6, tamanho - 1)
        End If
    End If
End Sub

Private Sub TextBox7_Change()
    If Len(TextBox7) > 0 Then
        tamanho = Len(TextBox7)
        ultimo = Right(TextBox7, 1)
    
        If Not IsNumeric(ultimo) Then
            TextBox7 = Left(TextBox7, tamanho - 1)
        End If
    End If
End Sub

Private Sub TextBox7_Enter()
    TextBox7.Text = ""
End Sub
