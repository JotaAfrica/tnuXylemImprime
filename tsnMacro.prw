


Sub tabItens()
    ' Cria as variáveis
    Dim Line As Integer
    Dim QtdeLines As Integer
    Dim Campo1 As String
    Dim Campo2 As String
    Dim Campo3 As String
 
    ' Posiciona no BookMark
    Selection.GoTo What:=wdGoToBookmark, Name:="tabItens"
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
 
    ' Busca a quantidade de linhas definida em QtdePro
    QtdeLines = Val(ActiveDocument.Variables.Item("QtdePro").Value)
     
    ' Percorre de 1 até o número de linhas definida
    For Line = 1 To QtdeLines
        ' Se for a coluna 1, preenche o valor com uma docvariable e pula para a coluna da direita
        Campo1 = "DOCVARIABLE ITEM1" & Trim(Str(Line))
        Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:=Campo1, PreserveFormatting:=True
        Selection.MoveRight Unit:=wdCell
         
        ' Se for a coluna 2, preenche o valor com uma docvariable e pula para a coluna da direita
        Campo2 = "DOCVARIABLE ITEM2" & Trim(Str(Line))
        Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:=Campo2, PreserveFormatting:=True
        Selection.MoveRight Unit:=wdCell
         
        ' Se for a coluna 3, preenche o valor com uma docvariable
        Campo3 = "DOCVARIABLE ITEM3" & Trim(Str(Line))
        Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:=Campo3, PreserveFormatting:=True
         
        ' Se a linha atual for menor que a última linha, pula para a direita (irá incluir uma nova linha)
        If Line < QtdeLines Then
            Selection.MoveRight Unit:=wdCell
        End If
    Next
End Sub
