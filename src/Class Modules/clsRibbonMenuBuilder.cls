Option Explicit

' =====================================================================
' CONFIGURAÇÃO INTERNA DA CLASSE
' =====================================================================
Private Const NOME_CONSULTA_PQ As String = "Variaveis" ' Nome da consulta do Power Query

' Variáveis de Estado (Ícones e Estruturas)
Private m_IconeOk As String
Private m_IconeAtencao As String
Private m_IconeErro As String
Private m_IconeCinza As String

Private m_SecoesOrdem As Collection
Private m_SecoesXML As Object
Private m_ConfigXML As String
Private m_idNum As Integer

' =====================================================================
' INICIALIZAÇÃO
' =====================================================================
Private Sub Class_Initialize()
    m_IconeOk = "NotebookColor12"       ' Verde
    m_IconeAtencao = "NotebookColor2"   ' Laranja/Amarelo
    m_IconeErro = "NotebookColor7"      ' Vermelho
    m_IconeCinza = "NotebookColor16"    ' Cinza
    
    Set m_SecoesOrdem = New Collection
    Set m_SecoesXML = CreateObject("Scripting.Dictionary")
    m_idNum = 1
    m_ConfigXML = ""
End Sub

' Propriedades para customização de ícones
Public Property Let IconeOk(valor As String): m_IconeOk = valor: End Property
Public Property Let IconeAtencao(valor As String): m_IconeAtencao = valor: End Property
Public Property Let IconeErro(valor As String): m_IconeErro = valor: End Property
Public Property Let IconeCinza(valor As String): m_IconeCinza = valor: End Property

' =====================================================================
' BLOCO 1: GERADORES DE XML (O Desenho do Menu)
' =====================================================================
Public Sub AdicionarMenuConfiguracao(DirAtual As String, NomeVariavelPQ As String)
    Dim xml As String
    xml = "<button id=""hdrConfig"" label=""CONFIGURAÇÃO"" enabled=""false""/>"
    xml = xml & "<menu id=""m_dirPadrao"" label=""Diretório Padrão"" image=""img_IrArquivo"" itemSize=""normal"">"
    xml = xml & "    <button id=""btnSelDirPadrao"" label=""Selecionar Diretório Padrão..."" image=""img_Buscar"" onAction=""OnAct_SelecionarPasta"" tag=""" & NomeVariavelPQ & "|" & DirAtual & """/>"
    xml = xml & "    <button id=""btnGoDirPadrao"" label=""Ir para Diretório Padrão"" image=""img_IrArquivo"" onAction=""OnAct_AbrirPasta"" tag=""" & DirAtual & """/>"
    xml = xml & "</menu>"
    m_ConfigXML = xml
End Sub

Public Sub CriarSecao(ID As String, Titulo As String, Optional IncluirSeparador As Boolean = True)
    Dim xmlCabecalho As String
    If m_SecoesXML.exists(ID) Then Exit Sub
    If IncluirSeparador Then xmlCabecalho = "<menuSeparator id=""sep_" & ID & """/>"
    xmlCabecalho = xmlCabecalho & "<button id=""hdr_" & ID & """ label=""" & Titulo & """ enabled=""false""/>"
    m_SecoesOrdem.Add ID
    m_SecoesXML.Add ID, xmlCabecalho
End Sub

Public Sub AdicionarArquivoModular(SecaoID As String, Nome As String, Prefixo As String, Extensao As String, DiretorioAlvo As String, Obrigatorio As Boolean, TagPQ As String, UsaPadrao As Boolean)
    Dim arquivoAtual As String, arquivoMaisRecente As String
    Dim dataAtual As Date, dataModificacao As Date
    Dim extReal As String, xml As String, caminho As String
    Dim corPai As String, lblPai As String, tagGo As String
    Dim lblUsarPadrao As String, btnSelecionarHabilitado As String, iconePadrao As String
    
    If Not m_SecoesXML.exists(SecaoID) Then Exit Sub
    
    ' Lógica do Toggle visual
    If UsaPadrao Then
        lblUsarPadrao = "Usar Padrão (Ativado)"
        iconePadrao = m_IconeOk
        btnSelecionarHabilitado = "false"
    Else
        lblUsarPadrao = "Usar Padrão (Desativado)"
        iconePadrao = m_IconeErro
        btnSelecionarHabilitado = "true"
    End If
    
    caminho = Trim(DiretorioAlvo)
    If caminho <> "" And Right(caminho, 1) <> "\" Then caminho = caminho & "\"
    arquivoMaisRecente = ""
    dataModificacao = 0
    
    ' Motor de Busca do Windows
    If caminho <> "" And Dir(caminho, vbDirectory) <> "" Then
        arquivoAtual = Dir(caminho & Prefixo & "*.*")
        Do While arquivoAtual <> ""
            If Left(arquivoAtual, 2) <> "~$" Then
                dataAtual = FileDateTime(caminho & arquivoAtual)
                If dataAtual > dataModificacao Then
                    dataModificacao = dataAtual
                    arquivoMaisRecente = arquivoAtual
                End If
            End If
            arquivoAtual = Dir()
        Loop
    End If
    
    If arquivoMaisRecente = "" Then
        lblPai = Nome & IIf(Obrigatorio, " (Faltando)", " (Não Encontrado)")
        corPai = IIf(Obrigatorio, m_IconeErro, m_IconeCinza)
        tagGo = caminho
        
        xml = "<menu id=""m_arq" & m_idNum & """ label=""" & lblPai & """ imageMso=""" & corPai & """ itemSize=""normal"">"
        xml = xml & "    <button id=""hdr_dir" & m_idNum & """ label=""DIRETÓRIO"" enabled=""false""/>"
        xml = xml & "    <button id=""btn_padrao" & m_idNum & """ label=""" & lblUsarPadrao & """ imageMso=""" & iconePadrao & """ onAction=""OnAct_UsarPadrao"" tag=""" & TagPQ & """/>"
        xml = xml & "    <button id=""btn_sel" & m_idNum & """ label=""Selecionar Nova Pasta..."" image=""img_Buscar"" enabled=""" & btnSelecionarHabilitado & """ onAction=""OnAct_SelecionarPasta"" tag=""" & TagPQ & "|" & caminho & """/>"
        xml = xml & "    <button id=""btn_go" & m_idNum & """ label=""Ir para Pasta"" image=""img_IrArquivo"" onAction=""OnAct_AbrirPasta"" tag=""" & tagGo & """/>"
        xml = xml & "</menu>"
    Else
        extReal = LCase(Mid(arquivoMaisRecente, InStrRev(arquivoMaisRecente, ".")))
        lblPai = Nome
        tagGo = caminho & arquivoMaisRecente
        
        Dim corExt As String, textoExt As String, corData As String, textoData As String
        corExt = IIf(extReal = LCase(Extensao), m_IconeOk, m_IconeErro)
        textoExt = "Extensão: " & extReal & IIf(extReal = LCase(Extensao), "", " (Incorreta)")
        corData = IIf(Int(dataModificacao) = Date, m_IconeOk, m_IconeAtencao)
        textoData = "Data: " & Format(dataModificacao, "dd/mm/yyyy hh:mm") & IIf(Int(dataModificacao) = Date, "", " (Antigo)")
        corPai = IIf(corExt = m_IconeErro Or corData = m_IconeAtencao, m_IconeAtencao, m_IconeOk)
        
        xml = "<menu id=""m_arq" & m_idNum & """ label=""" & lblPai & """ imageMso=""" & corPai & """ itemSize=""normal"">"
        xml = xml & "    <button id=""hdr_dir" & m_idNum & """ label=""DIRETÓRIO"" enabled=""false""/>"
        xml = xml & "    <button id=""btn_padrao" & m_idNum & """ label=""" & lblUsarPadrao & """ imageMso=""" & iconePadrao & """ onAction=""OnAct_UsarPadrao"" tag=""" & TagPQ & """/>"
        xml = xml & "    <button id=""btn_sel" & m_idNum & """ label=""Selecionar Nova Pasta..."" image=""img_Buscar"" enabled=""" & btnSelecionarHabilitado & """ onAction=""OnAct_SelecionarPasta"" tag=""" & TagPQ & "|" & caminho & """/>"
        xml = xml & "    <button id=""btn_go" & m_idNum & """ label=""Ir para Pasta"" image=""img_IrArquivo"" onAction=""OnAct_AbrirPasta"" tag=""" & tagGo & """/>"
        xml = xml & "    <menuSeparator id=""sep_int" & m_idNum & """/>"
        xml = xml & "    <button id=""hdr_info" & m_idNum & """ label=""INFORMAÇÕES"" enabled=""false""/>"
        xml = xml & "    <button id=""e_arq" & m_idNum & """ label=""" & textoExt & """ imageMso=""" & corExt & """ onAction=""MacroVazia""/>"
        xml = xml & "    <button id=""d_arq" & m_idNum & """ label=""" & textoData & """ imageMso=""" & corData & """ onAction=""MacroVazia""/>"
        xml = xml & "</menu>"
    End If
    
    m_SecoesXML(SecaoID) = m_SecoesXML(SecaoID) & xml
    m_idNum = m_idNum + 1
End Sub

Public Sub AdicionarArquivoDireto(SecaoID As String, Nome As String, CaminhoCompleto As String, Obrigatorio As Boolean, TagPQ As String)
    Dim xml As String, corPai As String, lblPai As String
    If Not m_SecoesXML.exists(SecaoID) Then Exit Sub
    
    If Trim(CaminhoCompleto) = "" Or Dir(CaminhoCompleto) = "" Then
        lblPai = Nome & IIf(Obrigatorio, " (Faltando)", " (Não Encontrado)")
        corPai = IIf(Obrigatorio, m_IconeErro, m_IconeCinza)
    Else
        lblPai = Nome
        corPai = m_IconeOk
    End If
    
    xml = "<menu id=""m_arq_dir" & m_idNum & """ label=""" & lblPai & """ imageMso=""" & corPai & """ itemSize=""normal"">"
    xml = xml & "    <button id=""hdr_dir_sys" & m_idNum & """ label=""DIRETÓRIO"" enabled=""false""/>"
    xml = xml & "    <button id=""btn_sel_sys" & m_idNum & """ label=""Selecionar Novo Arquivo..."" image=""img_Buscar"" onAction=""OnAct_SelecionarArquivo"" tag=""" & TagPQ & "|" & CaminhoCompleto & """/>"
    xml = xml & "    <button id=""btn_go_sys" & m_idNum & """ label=""Ir para Arquivo"" image=""img_IrArquivo"" onAction=""OnAct_AbrirPasta"" tag=""" & CaminhoCompleto & """/>"
    xml = xml & "</menu>"
    
    m_SecoesXML(SecaoID) = m_SecoesXML(SecaoID) & xml
    m_idNum = m_idNum + 1
End Sub

Public Function GerarXML() As String
    Dim xmlFinal As String, chaveSecao As Variant
    xmlFinal = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" itemSize=""large"">"
    If m_ConfigXML <> "" Then xmlFinal = xmlFinal & m_ConfigXML
    
    For Each chaveSecao In m_SecoesOrdem
        If InStr(1, m_SecoesXML(chaveSecao), "<menu ") > 0 Then xmlFinal = xmlFinal & m_SecoesXML(chaveSecao)
    Next chaveSecao
    
    GerarXML = xmlFinal & "</menu>"
End Function

' =====================================================================
' BLOCO 2: MOTOR DE AÇÕES (Lógica de Cliques e Windows)
' =====================================================================
Public Sub ExecutarTogglePadrao(Tag As String)
    Dim caminhoAtualPQ As String
    
    ' Tag contém apenas o nome da variável. Ex: "dirAderencia"
    caminhoAtualPQ = LerVariavelPQ_Interno(NOME_CONSULTA_PQ, Tag)
    
    If caminhoAtualPQ = "" Then
        ' Estava usando Padrão. O usuário clicou para desativar.
        ' Obrigamos ele a selecionar uma pasta para substituir o padrão.
        MsgBox "Para desativar o diretório padrão, por favor selecione a pasta específica que deseja utilizar.", vbInformation, "Desativar Padrão"
        ExecutarSelecionarPasta Tag & "|" ' Envia a Tag formatada com pipeline
    Else
        ' Não estava usando padrão. O usuário clicou para ativar o padrão.
        ' Apenas limpamos a variável.
        SalvarVariavelPQ_Interno NOME_CONSULTA_PQ, Tag, ""
    End If
End Sub

Public Sub ExecutarSelecionarPasta(Tag As String)
    Dim fd As FileDialog
    Dim partesTag() As String, nomeVariavel As String, caminhoInicial As String, novaPasta As String
    
    partesTag = Split(Tag, "|")
    If UBound(partesTag) < 0 Then Exit Sub
    nomeVariavel = partesTag(0)
    If UBound(partesTag) >= 1 Then caminhoInicial = partesTag(1) Else caminhoInicial = ""
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Selecione a pasta para: " & nomeVariavel
        If caminhoInicial <> "" And Dir(caminhoInicial, vbDirectory) <> "" Then .InitialFileName = caminhoInicial
        
        If .Show = -1 Then
            novaPasta = .SelectedItems(1)
            If Right(novaPasta, 1) <> "\" Then novaPasta = novaPasta & "\"
            SalvarVariavelPQ_Interno NOME_CONSULTA_PQ, nomeVariavel, novaPasta
        End If
    End With
End Sub

Public Sub ExecutarSelecionarArquivo(Tag As String)
    Dim fd As FileDialog
    Dim partesTag() As String, nomeVariavel As String, caminhoInicial As String, pastaInicial As String, novoArquivo As String
    
    partesTag = Split(Tag, "|")
    If UBound(partesTag) < 0 Then Exit Sub
    nomeVariavel = partesTag(0)
    If UBound(partesTag) >= 1 Then caminhoInicial = partesTag(1) Else caminhoInicial = ""
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Selecione o arquivo para: " & nomeVariavel
        If caminhoInicial <> "" Then
            pastaInicial = Left(caminhoInicial, InStrRev(caminhoInicial, "\"))
            If Dir(pastaInicial, vbDirectory) <> "" Then .InitialFileName = pastaInicial
        End If
        
        If .Show = -1 Then
            novoArquivo = .SelectedItems(1)
            SalvarVariavelPQ_Interno NOME_CONSULTA_PQ, nomeVariavel, novoArquivo
        End If
    End With
End Sub

Public Sub ExecutarAbrirCaminho(Tag As String)
    Dim pastaRaiz As String
    If Tag = "" Then Exit Sub
    
    On Error Resume Next
    If Dir(Tag, vbDirectory) <> "" Then
        If (GetAttr(Tag) And vbDirectory) = vbDirectory Then
            Shell "explorer.exe """ & Tag & """", vbNormalFocus ' É pasta
        Else
            Shell "explorer.exe /select,""" & Tag & """", vbNormalFocus ' É arquivo
        End If
    Else
        ' Se falhar, tenta abrir a pasta pai
        pastaRaiz = Left(Tag, InStrRev(Tag, "\"))
        If Dir(pastaRaiz, vbDirectory) <> "" Then Shell "explorer.exe """ & pastaRaiz & """", vbNormalFocus
    End If
    On Error GoTo 0
End Sub

' =====================================================================
' BLOCO 3: CRUD POWER QUERY (O "Banco de Dados" Embutido e Privado)
' =====================================================================
Private Sub SalvarVariavelPQ_Interno(nomeConsulta As String, nomeVariavel As String, valor As String)
    Dim valorAtual As String, valorFormatado As String
    
    PrepararConsultaRepositorio nomeConsulta
    valorAtual = LerVariavelPQ_Interno(nomeConsulta, nomeVariavel)
    valorFormatado = """" & valor & """" ' Coloca aspas para o M ler como Texto
    
    If valorAtual = "" Then
        AdicionarVariavelPQ_Interno nomeConsulta, nomeVariavel, valorFormatado
    Else
        AlterarVariavelPQ_Interno nomeConsulta, nomeVariavel, valorFormatado
    End If
End Sub

Private Function LerVariavelPQ_Interno(nomeConsulta As String, nomeVariavel As String) As String
    Dim formulaM As String, regex As Object, matches As Object
    LerVariavelPQ_Interno = ""
    
    On Error Resume Next
    formulaM = ActiveWorkbook.Queries(nomeConsulta).Formula
    On Error GoTo 0
    If formulaM = "" Then Exit Function

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\b" & nomeVariavel & "\s*=\s*([^,\r\n]+)"
    regex.IgnoreCase = True

    Set matches = regex.Execute(formulaM)
    If matches.Count > 0 Then
        LerVariavelPQ_Interno = Trim(matches(0).SubMatches(0))
        If Left(LerVariavelPQ_Interno, 1) = """" Then LerVariavelPQ_Interno = Mid(LerVariavelPQ_Interno, 2, Len(LerVariavelPQ_Interno) - 2)
    End If
End Function

Private Sub PrepararConsultaRepositorio(nomeConsulta As String)
    Dim qry As WorkbookQuery
    On Error Resume Next
    Set qry = ActiveWorkbook.Queries(nomeConsulta)
    On Error GoTo 0
    If qry Is Nothing Then ActiveWorkbook.Queries.Add Name:=nomeConsulta, Formula:="let" & vbCrLf & "    CriadoEm = """ & Now & """" & vbCrLf & "in" & vbCrLf & "    CriadoEm"
End Sub

Private Sub AdicionarVariavelPQ_Interno(nomeConsulta As String, nomeVariavel As String, valor As String)
    Dim formulaM As String, regex As Object
    formulaM = ActiveWorkbook.Queries(nomeConsulta).Formula
    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.Pattern = "\b(let)\b"
    If regex.Test(formulaM) Then
        formulaM = regex.Replace(formulaM, "$1" & vbCrLf & "    " & nomeVariavel & " = " & valor & ",")
        ActiveWorkbook.Queries(nomeConsulta).Formula = formulaM
    End If
End Sub

Private Sub AlterarVariavelPQ_Interno(nomeConsulta As String, nomeVariavel As String, novoValor As String)
    Dim formulaM As String, regex As Object
    formulaM = ActiveWorkbook.Queries(nomeConsulta).Formula
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "(\b" & nomeVariavel & "\s*=\s*)([^,\r\n]+)"
    regex.IgnoreCase = True
    If regex.Test(formulaM) Then
        formulaM = regex.Replace(formulaM, "$1" & novoValor)
        ActiveWorkbook.Queries(nomeConsulta).Formula = formulaM
    End If
End Sub