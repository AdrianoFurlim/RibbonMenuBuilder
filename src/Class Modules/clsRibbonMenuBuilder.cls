Option Explicit

' =====================================================================
' ENUMERAÇÕES
' =====================================================================
' ADICIONADO: Enumeração para o controle de tempo do arquivo (antigoEm)
Public Enum eValidade
    vLivre = 0
    v10Min = 1
    v30Min = 2
    v1Hora = 3
    v12Horas = 4
    v1Dia = 5
    v1Semana = 6
End Enum

' ADICIONADO: Enumeração para o tipo de cabeçalho da seção
Public Enum eTipoCabecalho
    eBtnSeparador = 0   ' Linha vazia + Botão Inativo (Padrão)
    ePdrSeparador = 1   ' Linha com o título embutido no próprio separador
End Enum

' =====================================================================
' CONFIGURAÇÃO INTERNA DA CLASSE
' =====================================================================
Private Const NOME_CONSULTA_PQ As String = "Variaveis"

Private m_IconeOk As String
Private m_IconeAtencao As String
Private m_IconeErro As String
Private m_IconeCinza As String

Private m_SecoesOrdem As Collection
Private m_SecoesXML As Object
Private m_idNum As Integer
Private m_MensagemValidacao As String

' ADICIONADO: Propriedade para controlar o tamanho do menu principal
Public TamanhoMenuPrincipal As String


' =====================================================================
' INICIALIZAÇÃO
' =====================================================================
Private Sub Class_Initialize()
    m_IconeOk = "NotebookColor12"
    m_IconeAtencao = "NotebookColor2"
    m_IconeErro = "NotebookColor7"
    m_IconeCinza = "NotebookColor16"
    m_MensagemValidacao = ""
    
    TamanhoMenuPrincipal = "large" ' Tamanho padrão inicial
    
    Set m_SecoesOrdem = New Collection
    Set m_SecoesXML = CreateObject("Scripting.Dictionary")
    m_idNum = 1
End Sub

Public Property Let IconeOk(valor As String): m_IconeOk = valor: End Property
Public Property Let IconeAtencao(valor As String): m_IconeAtencao = valor: End Property
Public Property Let IconeErro(valor As String): m_IconeErro = valor: End Property
Public Property Let IconeCinza(valor As String): m_IconeCinza = valor: End Property

' =====================================================================
' BLOCO 1: GERADORES DE XML (O Desenho do Menu)
' =====================================================================

Public Sub CriarSecao(ID As String, Titulo As String, Optional TipoCabecalho As eTipoCabecalho = eBtnSeparador)
    Dim xmlCabecalho As String
    
    ' Se a seção já existe, aborta para não duplicar
    If m_SecoesXML.exists(ID) Then Exit Sub
    
    ' Lógica de escolha do visual usando o Enum
    If TipoCabecalho = ePdrSeparador Then
        ' Usa o separador nativo que já possui o título embutido
        xmlCabecalho = "<menuSeparator id=""sep_" & ID & """ title=""" & Titulo & """/>"
    Else
        ' Fallback/Padrão (eBtnSeparador): Linha vazia seguida de um botão inativo
        xmlCabecalho = "<menuSeparator id=""sep_" & ID & """/>"
        xmlCabecalho = xmlCabecalho & "<button id=""hdr_" & ID & """ label=""" & Titulo & """ enabled=""false""/>"
    End If
    
    ' Registra a seção
    m_SecoesOrdem.Add ID
    m_SecoesXML.Add ID, xmlCabecalho
End Sub

' ADICIONADO: Método flexível para criar N diretórios padrões dentro de qualquer seção
Public Sub AdicionarDiretorioPadrao(SecaoID As String, Titulo As String, VarPQ As String, DirAtual As String, Optional TamanhoSubmenu As String = "normal")
    Dim xml As String
    If Not m_SecoesXML.exists(SecaoID) Then Exit Sub
    
    xml = "<menu id=""m_dir_" & m_idNum & """ label=""" & Titulo & """ image=""img_pasta"" itemSize=""" & TamanhoSubmenu & """>"
    xml = xml & "    <button id=""btnSelDir_" & m_idNum & """ label=""Selecionar Pasta Padrão..."" image=""img_procurar"" onAction=""OnAct_SelecionarPasta"" tag=""" & VarPQ & "|" & DirAtual & """/>"
    xml = xml & "    <button id=""btnGoDir_" & m_idNum & """ label=""Ir para Pasta Padrão"" image=""img_pastaAberta"" onAction=""OnAct_AbrirPasta"" tag=""" & DirAtual & """/>"
    xml = xml & "</menu>"
    
    m_SecoesXML(SecaoID) = m_SecoesXML(SecaoID) & xml
    m_idNum = m_idNum + 1
End Sub

' MODIFICADO: Assinatura atualizada (Prefixo -> NomeBusca, removido Obrigatorio, inserido AntigoEm e a lógica de Tags duplas)
Public Sub AdicionarArquivoModular(SecaoID As String, Nome As String, NomeBusca As String, Extensao As String, DiretorioAlvo As String, TagVarEspecifica As String, TagVarPadrao As String, ValorPadrao As String, Optional AntigoEm As eValidade = vLivre, Optional TamanhoSubmenu As String = "normal")
    Dim arquivoAtual As String, arquivoMaisRecente As String
    Dim dataAtual As Date, dataModificacao As Date
    Dim extReal As String, xml As String, caminho As String
    Dim corPai As String, lblPai As String, tagGo As String
    Dim lblUsarPadrao As String, btnSelecionarHabilitado As String, iconePadrao As String
    
    If Not m_SecoesXML.exists(SecaoID) Then Exit Sub
    
    ' LÓGICA DO TOGGLE: Compara o diretório alvo com a string atual do diretório padrão
    Dim isUsandoPadrao As Boolean
    isUsandoPadrao = (Trim(DiretorioAlvo) = Trim(ValorPadrao))
    
    If isUsandoPadrao Then
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
    
    If caminho <> "" And Dir(caminho, vbDirectory) <> "" Then
        arquivoAtual = Dir(caminho & NomeBusca & "*.*")
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
    
    ' A Tag do Toggle envia as DUAS variáveis para o PQ consultar e comparar depois
    Dim tagToggle As String
    tagToggle = TagVarEspecifica & "|" & TagVarPadrao
    
    If arquivoMaisRecente = "" Then
        lblPai = Nome & " (Não Encontrado)"
        corPai = m_IconeErro
        tagGo = caminho
        m_MensagemValidacao = m_MensagemValidacao & "> " & Nome & vbCrLf & "     • Não encontrado" & vbCrLf & vbCrLf
        
        xml = "<menu id=""m_arq" & m_idNum & """ label=""" & lblPai & """ imageMso=""" & corPai & """ itemSize=""" & TamanhoSubmenu & """>"
        xml = xml & "    <button id=""hdr_dir" & m_idNum & """ label=""DIRETÓRIO"" enabled=""false""/>"
        xml = xml & "    <button id=""btn_padrao" & m_idNum & """ label=""" & lblUsarPadrao & """ imageMso=""" & iconePadrao & """ onAction=""OnAct_UsarPadrao"" tag=""" & tagToggle & """/>"
        xml = xml & "    <button id=""btn_sel" & m_idNum & """ label=""Selecionar Nova Pasta..."" image=""img_procurar"" enabled=""" & btnSelecionarHabilitado & """ onAction=""OnAct_SelecionarPasta"" tag=""" & TagVarEspecifica & "|" & caminho & """/>"
        xml = xml & "    <button id=""btn_go" & m_idNum & """ label=""Ir para Pasta"" image=""img_pastaAberta"" onAction=""OnAct_AbrirPasta"" tag=""" & tagGo & """/>"
        xml = xml & "</menu>"
    Else
        extReal = LCase(Mid(arquivoMaisRecente, InStrRev(arquivoMaisRecente, ".")))
        lblPai = Nome
        tagGo = caminho & arquivoMaisRecente
        
        ' LÓGICA DO TEMPO (Enum)
        Dim isAntigo As Boolean
        isAntigo = False
        If AntigoEm <> vLivre Then
            Dim minsDiff As Long
            minsDiff = DateDiff("n", dataModificacao, Now)
            Select Case AntigoEm
                Case v10Min: isAntigo = (minsDiff > 10)
                Case v30Min: isAntigo = (minsDiff > 30)
                Case v1Hora: isAntigo = (minsDiff > 60)
                Case v12Horas: isAntigo = (minsDiff > 720)
                Case v1Dia: isAntigo = (minsDiff > 1440)
                Case v1Semana: isAntigo = (minsDiff > 10080)
            End Select
        End If
        
        Dim corExt As String, textoExt As String, corData As String, textoData As String
        corExt = IIf(extReal = LCase(Extensao), m_IconeOk, m_IconeErro)
        textoExt = "Extensão: " & extReal & IIf(extReal = LCase(Extensao), "", " (Incorreta)")
        
        corData = IIf(isAntigo, m_IconeAtencao, m_IconeOk)
        textoData = "Data: " & Format(dataModificacao, "dd/mm/yyyy hh:mm") & IIf(isAntigo, " (Antigo)", "")
        
        corPai = IIf(corExt = m_IconeErro Or corData = m_IconeAtencao, m_IconeAtencao, m_IconeOk)
        
        If extReal <> LCase(Extensao) Then
            m_MensagemValidacao = m_MensagemValidacao & "> " & Nome & vbCrLf & "     • Extensão incorreta (" & extReal & ")" & vbCrLf & vbCrLf
        ElseIf isAntigo Then
            m_MensagemValidacao = m_MensagemValidacao & "> " & Nome & vbCrLf & "     • Desatualizado: " & Format(dataModificacao, "dd/mm/yyyy hh:mm") & vbCrLf & vbCrLf
        End If

        xml = "<menu id=""m_arq" & m_idNum & """ label=""" & lblPai & """ imageMso=""" & corPai & """ itemSize=""" & TamanhoSubmenu & """>"
        xml = xml & "    <button id=""hdr_dir" & m_idNum & """ label=""DIRETÓRIO"" enabled=""false""/>"
        xml = xml & "    <button id=""btn_padrao" & m_idNum & """ label=""" & lblUsarPadrao & """ imageMso=""" & iconePadrao & """ onAction=""OnAct_UsarPadrao"" tag=""" & tagToggle & """/>"
        xml = xml & "    <button id=""btn_sel" & m_idNum & """ label=""Selecionar Nova Pasta..."" image=""img_procurar"" enabled=""" & btnSelecionarHabilitado & """ onAction=""OnAct_SelecionarPasta"" tag=""" & TagVarEspecifica & "|" & caminho & """/>"
        xml = xml & "    <button id=""btn_go" & m_idNum & """ label=""Ir para Pasta"" image=""img_pastaAberta"" onAction=""OnAct_AbrirPasta"" tag=""" & tagGo & """/>"
        xml = xml & "    <menuSeparator id=""sep_int" & m_idNum & """/>"
        xml = xml & "    <button id=""hdr_info" & m_idNum & """ label=""INFORMAÇÕES"" enabled=""false""/>"
        xml = xml & "    <button id=""e_arq" & m_idNum & """ label=""" & textoExt & """ imageMso=""" & corExt & """ onAction=""MacroVazia""/>"
        xml = xml & "    <button id=""d_arq" & m_idNum & """ label=""" & textoData & """ imageMso=""" & corData & """ onAction=""MacroVazia""/>"
        xml = xml & "</menu>"
    End If
    
    m_SecoesXML(SecaoID) = m_SecoesXML(SecaoID) & xml
    m_idNum = m_idNum + 1
End Sub

Public Sub AdicionarArquivoDireto(SecaoID As String, Nome As String, CaminhoCompleto As String, TagPQ As String, Optional TamanhoSubmenu As String = "normal")
    Dim xml As String, corPai As String, lblPai As String
    If Not m_SecoesXML.exists(SecaoID) Then Exit Sub
    
    If Trim(CaminhoCompleto) = "" Or Dir(CaminhoCompleto) = "" Then
        lblPai = Nome & " (Não Encontrado)"
        corPai = m_IconeErro
        m_MensagemValidacao = m_MensagemValidacao & "> " & Nome & vbCrLf & "     • Não encontrado ou inacessível" & vbCrLf & vbCrLf
    Else
        lblPai = Nome
        corPai = m_IconeOk
    End If
    
    xml = "<menu id=""m_arq_dir" & m_idNum & """ label=""" & lblPai & """ imageMso=""" & corPai & """ itemSize=""" & TamanhoSubmenu & """>"
    xml = xml & "    <button id=""hdr_dir_sys" & m_idNum & """ label=""DIRETÓRIO"" enabled=""false""/>"
    xml = xml & "    <button id=""btn_sel_sys" & m_idNum & """ label=""Selecionar Novo Arquivo..."" image=""img_procurar"" onAction=""OnAct_SelecionarArquivo"" tag=""" & TagPQ & "|" & CaminhoCompleto & """/>"
    xml = xml & "    <button id=""btn_go_sys" & m_idNum & """ label=""Ir para Arquivo"" image=""img_pastaAberta"" onAction=""OnAct_AbrirPasta"" tag=""" & CaminhoCompleto & """/>"
    xml = xml & "</menu>"
    
    m_SecoesXML(SecaoID) = m_SecoesXML(SecaoID) & xml
    m_idNum = m_idNum + 1
End Sub

Public Function GerarXML() As String
    Dim xmlFinal As String, chaveSecao As Variant
    xmlFinal = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" itemSize=""" & TamanhoMenuPrincipal & """>"
    
    For Each chaveSecao In m_SecoesOrdem
        If InStr(1, m_SecoesXML(chaveSecao), "<menu ") > 0 Then xmlFinal = xmlFinal & m_SecoesXML(chaveSecao)
    Next chaveSecao
    
    GerarXML = xmlFinal & "</menu>"
End Function

' =====================================================================
' BLOCO 2: MOTOR DE AÇÕES (Lógica de Cliques e Windows)
' =====================================================================
Public Sub ExecutarTogglePadrao(Tag As String)
    Dim partes() As String, varEspecifica As String, varPadrao As String
    Dim valEspecifico As String, valPadrao As String
    
    ' Separa as duas variáveis enviadas na Tag
    partes = Split(Tag, "|")
    If UBound(partes) < 1 Then Exit Sub
    varEspecifica = partes(0)
    varPadrao = partes(1)
    
    valEspecifico = LerVariavelPQ_Interno(NOME_CONSULTA_PQ, varEspecifica)
    valPadrao = LerVariavelPQ_Interno(NOME_CONSULTA_PQ, varPadrao)
    
    If valEspecifico = valPadrao Then
        ' Estão iguais. O usuário clicou para DESATIVAR o padrão.
        MsgBox "Para desativar o diretório padrão, por favor selecione a pasta específica que deseja utilizar.", vbInformation, "Desativar Padrão"
        ExecutarSelecionarPasta varEspecifica & "|" & valEspecifico
    Else
        ' Estão diferentes. O usuário clicou para ATIVAR o padrão.
        SalvarVariavelPQ_Interno NOME_CONSULTA_PQ, varEspecifica, valPadrao
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
            Shell "explorer.exe """ & Tag & """", vbNormalFocus
        Else
            Shell "explorer.exe /select,""" & Tag & """", vbNormalFocus
        End If
    Else
        pastaRaiz = Left(Tag, InStrRev(Tag, "\"))
        If Dir(pastaRaiz, vbDirectory) <> "" Then Shell "explorer.exe """ & pastaRaiz & """", vbNormalFocus
    End If
    On Error GoTo 0
End Sub

' =====================================================================
' BLOCO 3: CRUD POWER QUERY (O "Banco de Dados" Embutido e Privado)
' =====================================================================
Private Sub SalvarVariavelPQ_Interno(nomeConsulta As String, nomeVariavel As String, valor As String)
    Dim formulaM As String, regex As Object, valorFormatado As String
    
    PrepararConsultaRepositorio nomeConsulta
    formulaM = ActiveWorkbook.Queries(nomeConsulta).Formula
    valorFormatado = """" & valor & """"
    
    ' ALTERAÇÃO: Em vez de verificar se o "valor" é vazio,
    ' usamos o Regex para verificar se a declaração da variável já existe no código.
    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.Pattern = "\b" & nomeVariavel & "\s*=" ' Procura por: nomeVariavel =
    
    If regex.Test(formulaM) Then
        ' Se o nome foi encontrado, significa que a variável existe (mesmo se o valor for ""). Então ALTERA.
        AlterarVariavelPQ_Interno nomeConsulta, nomeVariavel, valorFormatado
    Else
        ' Se não encontrou o nome, ADICIONA.
        AdicionarVariavelPQ_Interno nomeConsulta, nomeVariavel, valorFormatado
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

' =====================================================================
' BLOCO 4: VALIDAÇÃO DE INTEGRIDADE (Trava Flexível)
' =====================================================================
Public Function ValidarArquivos() As Boolean
    ' Se a string não estiver vazia, significa que erros foram acumulados durante a criação
    If m_MensagemValidacao <> "" Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("Atenção! Foram encontradas pendências nos arquivos:" & vbCrLf & vbCrLf & _
                      m_MensagemValidacao & _
                      "Deseja prosseguir com a atualização do Power Query mesmo assim?", _
                      vbExclamation + vbYesNo + vbDefaultButton2, "Aviso de Validação de Arquivos")
                      
        If resp = vbNo Then
            ValidarArquivos = False
            Exit Function
        End If
    End If
    
    ' Se tudo estiver OK, ou se o usuário escolheu "SIM"
    ValidarArquivos = True
End Function


