V1:

Option Explicit   ' Obriga a declaração explícita de todas as variáveis, prevenindo erros de digitação ou uso indevido.

' ======================================================
' Função para normalizar texto
' ======================================================
Function NormalizarTexto(ByVal s As String) As String
    Dim temp As String

    ' Converte o texto para minúsculas e remove espaços extras nas extremidades
    temp = Trim(LCase(s))

    ' Remover pontuação final comum: .,;:!?)"'
    Do While Len(temp) > 0 And Right(temp, 1) Like "[.,;:!?)""']"
        temp = Trim(Left(temp, Len(temp) - 1))
    Loop

    ' Remover acentos e cedilha, substituindo por caracteres simples
    temp = Replace(temp, "á", "a")
    temp = Replace(temp, "à", "a")
    temp = Replace(temp, "ã", "a")
    temp = Replace(temp, "â", "a")
    temp = Replace(temp, "é", "e")
    temp = Replace(temp, "ê", "e")
    temp = Replace(temp, "í", "i")
    temp = Replace(temp, "ó", "o")
    temp = Replace(temp, "õ", "o")
    temp = Replace(temp, "ô", "o")
    temp = Replace(temp, "ú", "u")
    temp = Replace(temp, "ç", "c")

    ' Colapsar múltiplos espaços consecutivos em um único espaço
    Do While InStr(temp, "  ") > 0
        temp = Replace(temp, "  ", " ")
    Loop

    ' Retorna o texto normalizado
    NormalizarTexto = temp
End Function


' ======================================================
' Função para limpar prefixos comuns de e-mail: RE:, FW:, etc.
' ======================================================
Function LimparAssunto(ByVal s As String) As String
    Dim prefixos As Variant, p As Variant
    Dim mudou As Boolean

    ' Lista de prefixos a serem removidos, incluindo variações com e sem espaço
    prefixos = Array("RE:", "RES:", "FW:", "FWD:", "ENC:", "TR:", "RV:", "RESPOSTA:", _
                     "RE :", "FW :", "FWD :", "ENC :", "TR :", "RV :")

    s = Trim(s) ' Remove espaços extras nas extremidades

    ' Loop até que nenhum prefixo seja encontrado no início do texto
    Do
        mudou = False
        For Each p In prefixos
            If LCase(Left(s, Len(p))) = LCase(p) Then
                s = Trim(Mid(s, Len(p) + 1)) ' Remove o prefixo encontrado
                mudou = True
                Exit For
            End If
        Next p
    Loop While mudou

    ' Retorna o assunto sem prefixos
    LimparAssunto = s
End Function


' ======================================================
' Função para classificar o assunto do e-mail
' Inclui regras de exclusão e categorias
' ======================================================
Function ClassificarAssunto(ByVal assuntoOriginal As String) As String
    Dim n As String
    n = NormalizarTexto(assuntoOriginal)             ' Normaliza o texto

    ' Remove prefixos de e-mail e normaliza novamente
    n = NormalizarTexto(LimparAssunto(n))

    ' ==========================
    ' EXCLUSÃO — segunda coluna
    ' ==========================
    Dim excluir As Variant, ex As Variant
    excluir = Array( _
        "resumo diario de reacao", _
        "comunicados e noticias", _
        "registro de ponto", _
        "aviso sobre registro de solicitacao de ferias")

    ' Se o assunto for igual a algum da lista, marcar para exclusão
    For Each ex In excluir
        If n = ex Then
            ClassificarAssunto = "#EXCLUIR#"
            Exit Function
        End If
    Next ex

    ' ======================================================
    ' REGRA 1 – SIGAD
    ' ======================================================
    ' Verifica códigos SIGAD específicos
    If n Like "0100*" Or n Like "00100*" Or _
       n Like "0200*" Or n Like "00200*" Then
        ClassificarAssunto = "Ciência e acompanhamento no SIGAD"
        Exit Function
    End If

    ' Lista de termos que indicam acompanhamento no SIGAD
    Dim lista1 As Variant, t As Variant
    lista1 = Array("sem processo", "00100", "00200", _
                   "aviso de tramitacao", "aviso de distribuicao", _
                   "aviso de definicao de assinatura concluida")

    For Each t In lista1
        If InStr(n, t) > 0 Then
            ClassificarAssunto = "Ciência e acompanhamento no SIGAD"
            Exit Function
        End If
    Next t

    ' ======================================================
    ' REGRA 2 – Gestão Patrimonial
    ' ======================================================
    ' Lista de termos relacionados à Gestão Patrimonial
    Dim lista2 As Variant
    lista2 = Array("custo", _
                   "custos", _
                   "movimentacao direta", _
                   "movimentacao direta enviada por sinfra", _
                   "movimentacao direta recebida por sinfra")

    For Each t In lista2
        If InStr(n, t) > 0 Then
            ClassificarAssunto = "Gestão patrimonial"
            Exit Function
        End If
    Next t

    ' ======================================================
    ' PADRÃO
    ' ======================================================
    ' Caso não se encaixe em nenhuma regra anterior
    ClassificarAssunto = "Ciência e acompanhamento"
End Function


' ======================================================
' SUB PRINCIPAL: GerarPlanilha
' ======================================================
Sub GerarPlanilha()

    ' Declara objetos do Outlook e Excel
    Dim ns As Outlook.NameSpace
    Dim inbox As Outlook.Folder
    Dim inboxItems As Outlook.Folder
    Dim itm As Outlook.MailItem
    Dim dict As Object
    Dim key As Variant
    Dim subjectTxt As String
    Dim normalizado As String
    Dim filePath As String

    Dim xl As Object, wb As Object, ws As Object, wsAcoes As Object
    Dim row As Long

     ' Conecta ao Outlook
    Set ns = Application.GetNamespace("MAPI")
    Set inboxItems = ns.GetDefaultFolder(olFolderInbox)

    ' ======================================================
    ' TENTAR PEGAR A PASTA "resolvidos"
    ' ======================================================
    On Error Resume Next
    Set inbox = inboxItems.Folders("Resolvidos")
    On Error GoTo 0

    ' Se não existir, usar a Caixa de Entrada
    If inbox Is Nothing Then
        Set inbox = inboxItems
    End If

    Set dict = CreateObject("Scripting.Dictionary")

    ' ======================================================
    ' CAPTURAR ÚLTIMOS EMAILS POR ASSUNTO
    ' ======================================================
    Dim obj As Object
    For Each obj In inbox.Items
        If TypeOf obj Is Outlook.MailItem Then
            Set itm = obj

            subjectTxt = LimparAssunto(itm.Subject)
            normalizado = NormalizarTexto(subjectTxt)

            ' Adiciona ou substitui no dicionário apenas o último e-mail
            If normalizado <> "" Then
                If Not dict.Exists(normalizado) Then
                    dict.Add normalizado, itm
                ElseIf itm.ReceivedTime > dict(normalizado).ReceivedTime Then
                    Set dict(normalizado) = itm
                End If
            End If
        End If
    Next obj

    ' ======================================================
    ' ORDENAR POR DATA CRESCENTE + ASSUNTO
    ' ======================================================
    Dim arrKeys As Variant
    arrKeys = dict.Keys

    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(arrKeys) To UBound(arrKeys) - 1
        For j = i + 1 To UBound(arrKeys)
            Dim d1 As Date, d2 As Date
            d1 = Int(dict(arrKeys(i)).ReceivedTime)
            d2 = Int(dict(arrKeys(j)).ReceivedTime)

            ' Ordena por data, depois por assunto
            If (d1 > d2) Or (d1 = d2 And arrKeys(i) > arrKeys(j)) Then
                tmp = arrKeys(i)
                arrKeys(i) = arrKeys(j)
                arrKeys(j) = tmp
            End If
        Next j
    Next i

    ' ======================================================
    ' CRIAR PLANILHA EXCEL
    ' ======================================================
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False ' Mantém Excel oculto durante a execução

    Set wb = xl.Workbooks.Add
    Set ws = wb.Sheets(1)
    ws.Name = "Emails"

    ' Criar aba "Acoes"
    Set wsAcoes = wb.Sheets.Add(After:=ws)
    wsAcoes.Name = "Acoes"

    ' Lista de possíveis classificações / ações
    Dim lista As Collection: Set lista = New Collection
    lista.Add "Arquivamento"
    lista.Add "Autuação de processo"
    lista.Add "Cadastro no SIGAD"
    lista.Add "Cancelar tramitação"
    lista.Add "Ciência"
    lista.Add "Ciência e acompanhamento"
    lista.Add "Ciência e acompanhamento no SIGAD"
    lista.Add "Distribuição"
    lista.Add "Divulgação entre os colaboradores da SINFRA"
    lista.Add "Elaboração de documento"
    lista.Add "Encaminhamento aos líderes da SINFRA"
    lista.Add "Geração de vias"
    lista.Add "Gestão da unidade SINFRA no SIGAD"
    lista.Add "Gestão da caixa de e-mails da SINFRA"
    lista.Add "Gestão de pessoas"
    lista.Add "Gestão documental"
    lista.Add "Gestão junto à(s) área(s) responsável(eis) na SINFRA"
    lista.Add "Gestão junto à equipe do GBINFR"
    lista.Add "Gestão junto à equipe do GBINFR e ao Diretor da SINFRA"
    lista.Add "Gestão junto à(s) área(s) responsável(eis) externa(s) à SINFRA"
    lista.Add "Gestão junto à(s) área(s) envolvida(s)"
    lista.Add "Gestão junto aos Coordenadores"
    lista.Add "Gestão junto aos Coordenadores e Diretor da SINFRA"
    lista.Add "Gestão patrimonial"
    lista.Add "Juntada em processo"
    lista.Add "Juntada de documento"
    lista.Add "Orientações à(s) área(s)"
    lista.Add "Recebimento da Distribuição"
    lista.Add "Recebimento de tramitação"
    lista.Add "Retirada de documento"
    lista.Add "Tramitação"
    lista.Add "Tratativas com a Coordenação-Geral da SINFRA"
    lista.Add "Tratativas com a Diretoria e Coordenação-Geral da SINFRA"
    lista.Add "Tratativas com o Diretor da SINFRA"
    lista.Add "Tratativas com o suporte SIGAD"
    lista.Add "Treinamento Copilot"

    ' Cabeçalho da aba Acoes
    wsAcoes.Cells(1, 1).Value = "Ações"
    wsAcoes.Range("A1").Font.Bold = True
    wsAcoes.Range("A1").Font.Color = RGB(255, 255, 255)
    wsAcoes.Range("A1").Interior.Color = RGB(128, 128, 128)
    wsAcoes.Range("A1").HorizontalAlignment = -4108   ' Centralizar

    ' Preenche a lista de ações
    For i = 1 To lista.Count
        wsAcoes.Cells(i + 1, 1).Value = lista(i)
    Next i

    wsAcoes.Columns(1).AutoFit

    ' ======================================================
    ' Cabeçalho da planilha principal
    ' ======================================================
    ws.Cells(1, 1).Value = "DATA"
    ws.Cells(1, 2).Value = "NUP SIGAD / ASSUNTO"
    ws.Cells(1, 3).Value = "ATIVIDADE / ENTREGA"

    ws.Columns("B").NumberFormat = "@" ' Coluna B como texto

    With ws.Range("A1:C1")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 112, 192)
    End With
    
    ws.Range("A1:C1").HorizontalAlignment = -4108   ' Centralizar cabeçalhos

    ' ======================================================
    ' PREENCHER LINHAS COM EMAILS
    ' ======================================================
    row = 2
    Dim ultimaData As Date
    ultimaData = 0

    For Each key In arrKeys

        Set itm = dict(key)

        Dim classificacao As String
        classificacao = ClassificarAssunto(itm.Subject)

        ' Ignorar e-mails que devem ser excluídos
        If classificacao = "#EXCLUIR#" Then GoTo Continuar

        Dim dataAtual As Date
        dataAtual = Int(itm.ReceivedTime)

        ' Preenche a coluna de data apenas quando muda o dia
        If dataAtual <> ultimaData Then
            ws.Cells(row, 1).Value = dataAtual
            ws.Cells(row, 1).NumberFormat = "dd-mm-yyyy"
            ultimaData = dataAtual
        Else
            ws.Cells(row, 1).Value = ""
        End If

        ' Preenche assunto e classificação
        ws.Cells(row, 2).Value = LimparAssunto(itm.Subject)
        ws.Cells(row, 3).Value = classificacao

        row = row + 1

Continuar:
    Next key

    ' ======================================================
    ' VALIDATION NA COLUNA DE ATIVIDADES (C)
    ' ======================================================
    With ws.Columns("C").Validation
        .Delete
        .Add Type:=3, Formula1:="=Acoes!$A$2:$A$" & lista.Count + 1
    End With

    ' Ajustes finais
    ws.Columns.AutoFit
    ws.Columns("C").ColumnWidth = 40
    ws.Activate

    ' Salvar planilha na área de trabalho
    filePath = "D:\Users\andradej\OneDrive - Senado Federal\Área de Trabalho\PG Planilhas David\PG " & Format(Date, "dd-mm-yyyy") & ".xlsx"
    wb.SaveAs filePath
    wb.Close False
    xl.Quit

    ' Mensagem de conclusão
    MsgBox "Planilha criada: " & filePath, vbInformation

End Sub









VO:

Option Explicit   ' Obriga a declaração explícita de todas as variáveis, prevenindo erros de digitação ou uso indevido.

' ======================================================
' Função para normalizar texto
' ======================================================
Function NormalizarTexto(ByVal s As String) As String
    Dim temp As String

    ' Converte o texto para minúsculas e remove espaços extras nas extremidades
    temp = Trim(LCase(s))

    ' Remover pontuação final comum: .,;:!?)"'
    Do While Len(temp) > 0 And Right(temp, 1) Like "[.,;:!?)""']"
        temp = Trim(Left(temp, Len(temp) - 1))
    Loop

    ' Remover acentos e cedilha, substituindo por caracteres simples
    temp = Replace(temp, "á", "a")
    temp = Replace(temp, "à", "a")
    temp = Replace(temp, "ã", "a")
    temp = Replace(temp, "â", "a")
    temp = Replace(temp, "é", "e")
    temp = Replace(temp, "ê", "e")
    temp = Replace(temp, "í", "i")
    temp = Replace(temp, "ó", "o")
    temp = Replace(temp, "õ", "o")
    temp = Replace(temp, "ô", "o")
    temp = Replace(temp, "ú", "u")
    temp = Replace(temp, "ç", "c")

    ' Colapsar múltiplos espaços consecutivos em um único espaço
    Do While InStr(temp, "  ") > 0
        temp = Replace(temp, "  ", " ")
    Loop

    ' Retorna o texto normalizado
    NormalizarTexto = temp
End Function


' ======================================================
' Função para limpar prefixos comuns de e-mail: RE:, FW:, etc.
' ======================================================
Function LimparAssunto(ByVal s As String) As String
    Dim prefixos As Variant, p As Variant
    Dim mudou As Boolean

    ' Lista de prefixos a serem removidos, incluindo variações com e sem espaço
    prefixos = Array("RE:", "RES:", "FW:", "FWD:", "ENC:", "TR:", "RV:", "RESPOSTA:", _
                     "RE :", "FW :", "FWD :", "ENC :", "TR :", "RV :")

    s = Trim(s) ' Remove espaços extras nas extremidades

    ' Loop até que nenhum prefixo seja encontrado no início do texto
    Do
        mudou = False
        For Each p In prefixos
            If LCase(Left(s, Len(p))) = LCase(p) Then
                s = Trim(Mid(s, Len(p) + 1)) ' Remove o prefixo encontrado
                mudou = True
                Exit For
            End If
        Next p
    Loop While mudou

    ' Retorna o assunto sem prefixos
    LimparAssunto = s
End Function


' ======================================================
' Função para classificar o assunto do e-mail
' Inclui regras de exclusão e categorias
' ======================================================
Function ClassificarAssunto(ByVal assuntoOriginal As String) As String
    Dim n As String
    n = NormalizarTexto(assuntoOriginal)             ' Normaliza o texto

    ' Remove prefixos de e-mail e normaliza novamente
    n = NormalizarTexto(LimparAssunto(n))

    ' ==========================
    ' EXCLUSÃO — segunda coluna
    ' ==========================
    Dim excluir As Variant, ex As Variant
    excluir = Array( _
        "resumo diario de reacao", _
        "comunicados e noticias", _
        "registro de ponto", _
        "aviso sobre registro de solicitacao de ferias")

    ' Se o assunto for igual a algum da lista, marcar para exclusão
    For Each ex In excluir
        If n = ex Then
            ClassificarAssunto = "#EXCLUIR#"
            Exit Function
        End If
    Next ex

    ' ======================================================
    ' REGRA 1 – SIGAD
    ' ======================================================
    ' Verifica códigos SIGAD específicos
    If n Like "0100*" Or n Like "00100*" Or _
       n Like "0200*" Or n Like "00200*" Then
        ClassificarAssunto = "Ciência e acompanhamento no SIGAD"
        Exit Function
    End If

    ' Lista de termos que indicam acompanhamento no SIGAD
    Dim lista1 As Variant, t As Variant
    lista1 = Array("sem processo", "00100", "00200", _
                   "aviso de tramitacao", "aviso de distribuicao", _
                   "aviso de definicao de assinatura concluida")

    For Each t In lista1
        If InStr(n, t) > 0 Then
            ClassificarAssunto = "Ciência e acompanhamento no SIGAD"
            Exit Function
        End If
    Next t

    ' ======================================================
    ' REGRA 2 – Gestão Patrimonial
    ' ======================================================
    ' Lista de termos relacionados à Gestão Patrimonial
    Dim lista2 As Variant
    lista2 = Array("custo final", _
                   "movimentacao direta enviada por sinfra", _
                   "custo inicial")

    For Each t In lista2
        If InStr(n, t) > 0 Then
            ClassificarAssunto = "Gestão patrimonial"
            Exit Function
        End If
    Next t

    ' ======================================================
    ' PADRÃO
    ' ======================================================
    ' Caso não se encaixe em nenhuma regra anterior
    ClassificarAssunto = "Ciência e acompanhamento"
End Function


' ======================================================
' SUB PRINCIPAL: GerarPlanilha
' ======================================================
Sub GerarPlanilha()

    ' Declara objetos do Outlook e Excel
    Dim ns As Outlook.NameSpace
    Dim inbox As Outlook.Folder
    Dim inboxItems As Outlook.Folder
    Dim itm As Outlook.MailItem
    Dim dict As Object
    Dim key As Variant
    Dim subjectTxt As String
    Dim normalizado As String
    Dim filePath As String

    Dim xl As Object, wb As Object, ws As Object, wsAcoes As Object
    Dim row As Long

     ' Conecta ao Outlook
    Set ns = Application.GetNamespace("MAPI")
    Set inboxItems = ns.GetDefaultFolder(olFolderInbox)

    ' ======================================================
    ' TENTAR PEGAR A PASTA "resolvidos"
    ' ======================================================
    On Error Resume Next
    Set inbox = inboxItems.Folders("Resolvidos")
    On Error GoTo 0

    ' Se não existir, usar a Caixa de Entrada
    If inbox Is Nothing Then
        Set inbox = inboxItems
    End If

    Set dict = CreateObject("Scripting.Dictionary")

    ' ======================================================
    ' CAPTURAR ÚLTIMOS EMAILS POR ASSUNTO
    ' ======================================================
    Dim obj As Object
    For Each obj In inbox.Items
        If TypeOf obj Is Outlook.MailItem Then
            Set itm = obj

            subjectTxt = LimparAssunto(itm.Subject)
            normalizado = NormalizarTexto(subjectTxt)

            ' Adiciona ou substitui no dicionário apenas o último e-mail
            If normalizado <> "" Then
                If Not dict.Exists(normalizado) Then
                    dict.Add normalizado, itm
                ElseIf itm.ReceivedTime > dict(normalizado).ReceivedTime Then
                    Set dict(normalizado) = itm
                End If
            End If
        End If
    Next obj

    ' ======================================================
    ' ORDENAR POR DATA CRESCENTE + ASSUNTO
    ' ======================================================
    Dim arrKeys As Variant
    arrKeys = dict.Keys

    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(arrKeys) To UBound(arrKeys) - 1
        For j = i + 1 To UBound(arrKeys)
            Dim d1 As Date, d2 As Date
            d1 = Int(dict(arrKeys(i)).ReceivedTime)
            d2 = Int(dict(arrKeys(j)).ReceivedTime)

            ' Ordena por data, depois por assunto
            If (d1 > d2) Or (d1 = d2 And arrKeys(i) > arrKeys(j)) Then
                tmp = arrKeys(i)
                arrKeys(i) = arrKeys(j)
                arrKeys(j) = tmp
            End If
        Next j
    Next i

    ' ======================================================
    ' CRIAR PLANILHA EXCEL
    ' ======================================================
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False ' Mantém Excel oculto durante a execução

    Set wb = xl.Workbooks.Add
    Set ws = wb.Sheets(1)
    ws.Name = "Emails"

    ' Criar aba "Acoes"
    Set wsAcoes = wb.Sheets.Add(After:=ws)
    wsAcoes.Name = "Acoes"

    ' Lista de possíveis classificações / ações
    Dim lista As Collection: Set lista = New Collection
    lista.Add "Arquivamento"
    lista.Add "Autuação de processo"
    lista.Add "Cadastro no SIGAD"
    lista.Add "Cancelar tramitação"
    lista.Add "Ciência"
    lista.Add "Ciência e acompanhamento"
    lista.Add "Ciência e acompanhamento no SIGAD"
    lista.Add "Distribuição"
    lista.Add "Divulgação entre os colaboradores da SINFRA"
    lista.Add "Elaboração de documento"
    lista.Add "Encaminhamento aos líderes da SINFRA"
    lista.Add "Geração de vias"
    lista.Add "Gestão da unidade SINFRA no SIGAD"
    lista.Add "Gestão da caixa de e-mails da SINFRA"
    lista.Add "Gestão de pessoas"
    lista.Add "Gestão documental"
    lista.Add "Gestão junto à(s) área(s) responsável(eis) na SINFRA"
    lista.Add "Gestão junto à equipe do GBINFR"
    lista.Add "Gestão junto à equipe do GBINFR e ao Diretor da SINFRA"
    lista.Add "Gestão junto à(s) área(s) responsável(eis) externa(s) à SINFRA"
    lista.Add "Gestão junto à(s) área(s) envolvida(s)"
    lista.Add "Gestão junto aos Coordenadores"
    lista.Add "Gestão junto aos Coordenadores e Diretor da SINFRA"
    lista.Add "Gestão patrimonial"
    lista.Add "Juntada em processo"
    lista.Add "Juntada de documento"
    lista.Add "Orientações à(s) área(s)"
    lista.Add "Recebimento da Distribuição"
    lista.Add "Recebimento de tramitação"
    lista.Add "Retirada de documento"
    lista.Add "Tramitação"
    lista.Add "Tratativas com a Coordenação-Geral da SINFRA"
    lista.Add "Tratativas com a Diretoria e Coordenação-Geral da SINFRA"
    lista.Add "Tratativas com o Diretor da SINFRA"
    lista.Add "Tratativas com o suporte SIGAD"
    lista.Add "Treinamento Copilot"

    ' Cabeçalho da aba Acoes
    wsAcoes.Cells(1, 1).Value = "Ações"
    wsAcoes.Range("A1").Font.Bold = True
    wsAcoes.Range("A1").Font.Color = RGB(255, 255, 255)
    wsAcoes.Range("A1").Interior.Color = RGB(128, 128, 128)
    wsAcoes.Range("A1").HorizontalAlignment = -4108   ' Centralizar

    ' Preenche a lista de ações
    For i = 1 To lista.Count
        wsAcoes.Cells(i + 1, 1).Value = lista(i)
    Next i

    wsAcoes.Columns(1).AutoFit

    ' ======================================================
    ' Cabeçalho da planilha principal
    ' ======================================================
    ws.Cells(1, 1).Value = "DATA"
    ws.Cells(1, 2).Value = "NUP SIGAD / ASSUNTO"
    ws.Cells(1, 3).Value = "ATIVIDADE / ENTREGA"

    ws.Columns("B").NumberFormat = "@" ' Coluna B como texto

    With ws.Range("A1:C1")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 112, 192)
    End With
    
    ws.Range("A1:C1").HorizontalAlignment = -4108   ' Centralizar cabeçalhos

    ' ======================================================
    ' PREENCHER LINHAS COM EMAILS
    ' ======================================================
    row = 2
    Dim ultimaData As Date
    ultimaData = 0

    For Each key In arrKeys

        Set itm = dict(key)

        Dim classificacao As String
        classificacao = ClassificarAssunto(itm.Subject)

        ' Ignorar e-mails que devem ser excluídos
        If classificacao = "#EXCLUIR#" Then GoTo Continuar

        Dim dataAtual As Date
        dataAtual = Int(itm.ReceivedTime)

        ' Preenche a coluna de data apenas quando muda o dia
        If dataAtual <> ultimaData Then
            ws.Cells(row, 1).Value = dataAtual
            ws.Cells(row, 1).NumberFormat = "dd-mm-yyyy"
            ultimaData = dataAtual
        Else
            ws.Cells(row, 1).Value = ""
        End If

        ' Preenche assunto e classificação
        ws.Cells(row, 2).Value = LimparAssunto(itm.Subject)
        ws.Cells(row, 3).Value = classificacao

        row = row + 1

Continuar:
    Next key

    ' ======================================================
    ' VALIDATION NA COLUNA DE ATIVIDADES (C)
    ' ======================================================
    With ws.Columns("C").Validation
        .Delete
        .Add Type:=3, Formula1:="=Acoes!$A$2:$A$" & lista.Count + 1
    End With

    ' Ajustes finais
    ws.Columns.AutoFit
    ws.Columns("C").ColumnWidth = 40
    ws.Activate

    ' Salvar planilha na área de trabalho
    filePath = "D:\Users\andradej\OneDrive - Senado Federal\Área de Trabalho\PG Planilhas David\PG " & Format(Date, "dd-mm-yyyy") & ".xlsx"
    wb.SaveAs filePath
    wb.Close False
    xl.Quit

    ' Mensagem de conclusão
    MsgBox "Planilha criada: " & filePath, vbInformation

End Sub

                  
