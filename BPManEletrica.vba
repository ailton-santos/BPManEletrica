Option Explicit

' Declarando as variáveis
Dim faturamento As Double ' Faturamento mensal da empresa
Dim despesasFixas As Double ' Despesas fixas mensais da empresa
Dim custosVariaveis As Double ' Custos variáveis por atendimento
Dim qtdAtendimentos As Integer ' Quantidade de atendimentos mensais

' Função para calcular o lucro líquido mensal
Function lucroLiquido() As Double
    lucroLiquido = faturamento - despesasFixas - (custosVariaveis * qtdAtendimentos)
End Function

' Função para mostrar o Business Plan
Sub mostrarBusinessPlan()
    ' Limpando a planilha
    Sheets("Business Plan").Cells.ClearContents
    
    ' Configurando a planilha
    With Sheets("Business Plan")
        .Range("A1").Value = "Business Plan"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        
        .Range("A3").Value = "Faturamento mensal:"
        .Range("B3").Value = Format(faturamento, "R$ #,##0.00")
        
        .Range("A4").Value = "Despesas fixas mensais:"
        .Range("B4").Value = Format(despesasFixas, "R$ #,##0.00")
        
        .Range("A5").Value = "Custos variáveis por atendimento:"
        .Range("B5").Value = Format(custosVariaveis, "R$ #,##0.00")
        
        .Range("A6").Value = "Quantidade de atendimentos mensais:"
        .Range("B6").Value = qtdAtendimentos
        
        .Range("A8").Value = "Lucro líquido mensal:"
        .Range("B8").Value = Format(lucroLiquido(), "R$ #,##0.00")
        If lucroLiquido() > 0 Then
            .Range("B8").Font.Color = vbGreen ' Se o lucro for positivo, mostra em verde
        Else
            .Range("B8").Font.Color = vbRed ' Se o lucro for negativo, mostra em vermelho
        End If
    End With
End Sub

' Função para receber os valores das variáveis
Sub receberValores()
    faturamento = InputBox("Digite o faturamento mensal da empresa:")
    despesasFixas = InputBox("Digite as despesas fixas mensais da empresa:")
    custosVariaveis = InputBox("Digite os custos variáveis por atendimento:")
    qtdAtendimentos = InputBox("Digite a quantidade de atendimentos mensais:")
End Sub

' Função principal
Sub main()
    receberValores() ' Recebe os valores das variáveis
    mostrarBusinessPlan() ' Mostra o Business Plan
End Sub
