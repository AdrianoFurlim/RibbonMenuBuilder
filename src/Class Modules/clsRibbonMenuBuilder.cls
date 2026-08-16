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
' ==========================================================================
' Rotina:      Class_Initialize
' Descrição:   Construtor da classe. Inicializa as variáveis de estado, 
'              ícones padrão e os objetos de coleção para o controle do XML.
' Parâmetros:  Nenhum
' ==========================================================================
Private Sub Class_Initialize()
    ' Definição dos ícones nativos do Office (imageMso) para status visual
    m_IconeOk = "NotebookColor12"
    m_IconeAtencao = "NotebookColor2"
    m_IconeErro = "NotebookColor7"
    m_IconeCinza = "NotebookColor16"
    m_MensagemValidacao = ""
    
    TamanhoMenuPrincipal = "large" ' Tamanho padrão inicial
    
    ' Inicializa as coleções que armazenarão a estrutura do menu em memória
    Set m_SecoesOrdem = New Collection
    Set m_SecoesXML = CreateObject("Scripting.Dictionary")
    m_idNum = 1
End Sub

' Propriedades (Getters/Setters) para customização externa dos ícones
Public Property Let IconeOk(valor As String): m_IconeOk = valor: End Property
Public Property Let IconeAtencao(valor As String): m_IconeAtencao = valor: End Property
Public Property Let IconeErro(valor As String): m_IconeErro = valor: End Property
Public Property Let IconeCinza(valor As String): m_IconeCinza = valor: End Property

' =====================================================================
' BLOCO 1: GERADORES DE XML (O Desenho do Menu)
' =====================================================================

' ==========================================================================
' Rotina:      CriarSecao
' Descrição:   Cria um novo agrupamento (seção) no menu CustomUI.
' Parâmetros:  ID (String) - Identificador único da seção.
'              Titulo (String) - Texto exibido no cabeçalho.
'              TipoCabecalho (eTipoCabecalho) - Estilo visual do separador.
' ==========================================================================
Public Sub CriarSecao(ID As String, Titulo As String, Optional TipoCabecalho As eTipoCabecalho = ePdrSeparador)
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
    
    ' Registra a seção mantendo a ordem de inserção e armazenando o XML inicial
    m_SecoesOrdem.Add ID
    m_SecoesXML.Add ID, xmlCabecalho
End Sub

' ==========================================================================
' Rotina:      AdicionarDiretorioPadrao
' Descrição:   Cria botões genéricos para seleção e navegação de diretórios.
' Parâmetros:  SecaoID (String) - ID da seção pai.
'              Titulo (String) - Nome exibido no menu.
'              VarPQ (String) - Nome da variável no Power Query.
'              DirAtual (String) - Caminho atual configurado.
'              TamanhoSubmenu (String) - Define se o item é normal ou largo.
' ==========================================================================
' ADICIONADO: Método flexível para criar N diretórios padrões dentro de qualquer seção
Public Sub AdicionarDiretorioPadrao(SecaoID As String, Titulo As String, VarPQ As String, DirAtual As String, Optional TamanhoSubmenu As String = "normal")
    Dim xml As String
    ' Verifica a integridade garantindo que a seção pai existe
    If Not m_SecoesXML.exists(SecaoID) Then Exit Sub
    
    ' Constrói a estrutura XML do subset de diretório
    xml = "<menu id=""m_dir_" & m_idNum & """ label=""" & Titulo & """ image=""img_pasta"" itemSize=""" & TamanhoSubmenu & """>"
    xml = xml & "    <button id=""btnSelDir_" & m_idNum & """ label=""Selecionar Pasta Padrão..."" image=""img_procurar"" onAction=""OnAct_SelecionarPasta"" tag=""" & VarPQ & "|" & DirAtual & """/>"
    xml = xml & "    <button id=""btnGoDir_" & m_idNum & """ label=""Ir para Pasta Padrão"" image=""img_pastaAberta"" onAction=""OnAct_AbrirPasta"" tag=""" & DirAtual & """/>"
    xml = xml & "</menu>"
    
    ' Concatena o novo XML na seção correspondente e incrementa o ID global
    m_SecoesXML(SecaoID) = m_SecoesXML(SecaoID) & xml
    m_idNum = m_idNum + 1
End Sub

' ==========================================================================
' Rotina:      AdicionarArquivoModular
' Descrição:   Verifica a existência, extensão e validade temporal de um 
'              arquivo específico, gerando o XML correspondente com alertas 
'              visuais caso haja inconsistências.
' Parâmetros:  SecaoID, Nome, NomeBusca, Extensao, DiretorioAlvo, TagVarEspecifica,
'              TagVarPadrao, ValorPadrao, AntigoEm, TamanhoSubmenu, TipoCabecalhoSubmenu.
' ==========================================================================
' MODIFICADO: Assinatura atualizada (Prefixo -> NomeBusca, removido Obrigatorio, inserido AntigoEm e a lógica de Tags duplas)
Public Sub AdicionarArquivoModular(SecaoID As String, Nome As String, NomeBusca As String, Extensao As String, DiretorioAlvo As String, TagVarEspecifica As String, TagVarPadrao As String, ValorPadrao As String, Optional AntigoEm As eValidade = vLivre, Optional TamanhoSubmenu As String = "normal", Optional TipoCabecalhoSubmenu As eTipoCabecalho = ePdrSeparador)
    Dim arquivoAtual As String, arquivoMaisRecente As String
    Dim dataAtual As Date, dataModificacao As Date
    Dim extReal As String, xml As String, caminho As String
    Dim corPai As String, lblPai As String, tagGo As String
    Dim lblUsarPadrao As String, btnSelecionarHabilitado As String, iconePadrao As String
    
    If Not m_SecoesXML.exists(SecaoID) Then Exit Sub
    
    ' LÓGICA DO TOGGLE: Compara o diretório alvo com a string atual do diretório padrão
    Dim isUsandoPadrao As Boolean
    isUsandoPadrao = (Trim(DiretorioAlvo) = Trim(ValorPadrao))
    
    ' Define os estados visuais do botão Toggle baseando-se na correspondência com o padrão
    If isUsandoPadrao Then
        lblUsarPadrao = "Usar Padrão (Ativado)"
        iconePadrao = m_IconeOk
        btnSelecionarHabilitado = "false"
    Else
        lblUsarPadrao = "Usar Padrão (Desativado)"
        iconePadrao = m_IconeErro
        btnSelecionarHabilitado = "true"
    End If
    
    ' Garante que o caminho do diretório termina com a barra invertida apropriada
    caminho = Trim(DiretorioAlvo)
    If caminho <> "" And Right(caminho, 1) <> "\" Then caminho = caminho & "\"
    arquivoMaisRecente = ""
    dataModificacao = 0
    
    ' Varre o diretório em busca do arquivo mais recente que corresponda ao padrão de busca
    If caminho <> "" And Dir(caminho, vbDirectory) <> "" Then
        arquivoAtual = Dir(caminho & NomeBusca & "*.*")
        Do While arquivoAtual <> ""
            ' Ignora arquivos temporários do sistema (iniciados com ~$)
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
    
    ' Fluxo Alternativo: O arquivo não foi encontrado no diretório mapeado
    If arquivoMaisRecente = "" Then
        lblPai = Nome & " (Não Encontrado)"
        corPai = m_IconeErro
        tagGo = caminho
        
        ' Acumula a falha na string de validação global
        m_MensagemValidacao = m_MensagemValidacao & "> " & Nome & vbCrLf & "     • Não encontrado" & vbCrLf & vbCrLf
        
        xml = "<menu id=""m_arq" & m_idNum & """ label=""" & lblPai & """ imageMso=""" & corPai & """ itemSize=""" & TamanhoSubmenu & """>"
        
        ' --- CONTROLE DO TÍTULO: DIRETÓRIO ---
        If TipoCabecalhoSubmenu = ePdrSeparador Then
            xml = xml & "    <menuSeparator id=""sep_dir" & m_idNum & """ title=""DIRETÓRIO""/>"
        Else
            xml = xml & "    <button id=""hdr_dir" & m_idNum & """ label=""DIRETÓRIO"" enabled=""false""/>"
        End If
        
        ' Montagem dos controles para situação de arquivo não encontrado
        xml = xml & "    <button id=""btn_padrao" & m_idNum & """ label=""" & lblUsarPadrao & """ imageMso=""" & iconePadrao & """ onAction=""OnAct_UsarPadrao"" tag=""" & tagToggle & """/>"
        xml = xml & "    <button id=""btn_sel" & m_idNum & """ label=""Selecionar Nova Pasta..."" image=""img_procurar"" enabled=""" & btnSelecionarHabilitado & """ onAction=""OnAct_SelecionarPasta"" tag=""" & TagVarEspecifica & "|" & caminho & """/>"
        xml = xml & "    <button id=""btn_go" & m_idNum & """ label=""Ir para Pasta"" image=""img_pastaAberta"" onAction=""OnAct_AbrirPasta"" tag=""" & tagGo & """/>"
        xml = xml & "</menu>"
    Else
        ' Fluxo Principal: O arquivo foi encontrado, realiza as validações de regra de negócio
        extReal = LCase(Mid(arquivoMaisRecente, InStrRev(arquivoMaisRecente, ".")))
        lblPai = Nome
        tagGo = caminho & arquivoMaisRecente
        
        ' Validação Temporal: Verifica se o arquivo está obsoleto conforme o Enum configurado
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
        
        ' Define dinamicamente as cores de feedback baseadas nas validações
        Dim corExt As String, textoExt As String, corData As String, textoData As String
        corExt = IIf(extReal = LCase(Extensao), m_IconeOk, m_IconeErro)
        textoExt = "Extensão: " & extReal & IIf(extReal = LCase(Extensao), "", " (Incorreta)")
        
        corData = IIf(isAntigo, m_IconeAtencao, m_IconeOk)
        textoData = "Data: " & Format(dataModificacao, "dd/mm/yyyy hh:mm") & IIf(isAntigo, " (Antigo)", "")
        
        ' Consolida o status do botão pai (se há erro ou aviso nas validações)
        corPai = IIf(corExt = m_IconeErro Or corData = m_IconeAtencao, m_IconeAtencao, m_IconeOk)
        
        ' Registra inconsistências para o prompt de validação final
        If extReal <> LCase(Extensao) Then
            m_MensagemValidacao = m_MensagemValidacao & "> " & Nome & vbCrLf & "     • Extensão incorreta (" & extReal & ")" & vbCrLf & vbCrLf
        ElseIf isAntigo Then
            m_MensagemValidacao = m_MensagemValidacao & "> " & Nome & vbCrLf & "     • Desatualizado: " & Format(dataModificacao, "dd/mm/yyyy hh:mm") & vbCrLf & vbCrLf
        End If
        
        xml = "<menu id=""m_arq" & m_idNum & """ label=""" & lblPai & """ imageMso=""" & corPai & """ itemSize=""" & TamanhoSubmenu & """>"
        
        ' --- CONTROLE DO TÍTULO: DIRETÓRIO ---
        If TipoCabecalhoSubmenu = ePdrSeparador Then
            xml = xml & "    <menuSeparator id=""sep_dir" & m_idNum & """ title=""DIRETÓRIO""/>"
        Else
            xml = xml & "    <button id=""hdr_dir" & m_idNum & """ label=""DIRETÓRIO"" enabled=""false""/>"
        End If
        
        xml = xml & "    <button id=""btn_padrao" & m_idNum & """ label=""" & lblUsarPadrao & """ imageMso=""" & iconePadrao & """ onAction=""OnAct_UsarPadrao"" tag=""" & tagToggle & """/>"
        xml = xml & "    <button id=""btn_sel" & m_idNum & """ label=""Selecionar Nova Pasta..."" image=""img_procurar"" enabled=""" & btnSelecionarHabilitado & """ onAction=""OnAct_SelecionarPasta"" tag=""" & TagVarEspecifica & "|" & caminho & """/>"
        xml = xml & "    <button id=""btn_go" & m_idNum & """ label=""Ir para Pasta"" image=""img_pastaAberta"" onAction=""OnAct_AbrirPasta"" tag=""" & tagGo & """/>"
        
        ' --- CONTROLE DO TÍTULO: INFORMAÇÕES ---
        If TipoCabecalhoSubmenu = ePdrSeparador Then
            xml = xml & "    <menuSeparator id=""sep_info" & m_idNum & """ title=""INFORMAÇÕES""/>"
        Else
            xml = xml & "    <menuSeparator id=""sep_int" & m_idNum & """/>"
            xml = xml & "    <button id=""hdr_info" & m_idNum & """ label=""INFORMAÇÕES"" enabled=""false""/>"
        End If
        
        ' Inclusão de botões informativos (sem ação) exibindo os dados de validação
        xml = xml & "    <button id=""e_arq" & m_idNum & """ label=""" & textoExt & """ imageMso=""" & corExt & """ onAction=""MacroVazia""/>"
        xml = xml & "    <button id=""d_arq" & m_idNum & """ label=""" & textoData & """ imageMso=""" & corData & """ onAction=""MacroVazia""/>"
        xml = xml & "</menu>"
    End If
    
    m_SecoesXML(SecaoID) = m_SecoesXML(SecaoID) & xml
    m_idNum = m_idNum + 1
End Sub

' ==========================================================================
' Rotina:      AdicionarArquivoDireto
' Descrição:   Cria a interface para um arquivo específico apontado diretamente 
'              pelo usuário (sem busca em diretório).
' Parâmetros:  SecaoID, Nome, CaminhoCompleto, TagPQ, TamanhoSubmenu, TipoCabecalhoSubmenu
' ==========================================================================
Public Sub AdicionarArquivoDireto(SecaoID As String, Nome As String, CaminhoCompleto As String, TagPQ As String, Optional TamanhoSubmenu As String = "normal", Optional TipoCabecalhoSubmenu As eTipoCabecalho = ePdrSeparador)
    Dim xml As String, corPai As String, lblPai As String
    If Not m_SecoesXML.exists(SecaoID) Then Exit Sub
    
    ' Valida o acesso direto ao arquivo no sistema operacional
    If Trim(CaminhoCompleto) = "" Or Dir(CaminhoCompleto) = "" Then
        lblPai = Nome & " (Não Encontrado)"
        corPai = m_IconeErro
        m_MensagemValidacao = m_MensagemValidacao & "> " & Nome & vbCrLf & "     • Não encontrado ou inacessível" & vbCrLf & vbCrLf
    Else
        lblPai = Nome
        corPai = m_IconeOk
    End If
    
    xml = "<menu id=""m_arq_dir" & m_idNum & """ label=""" & lblPai & """ imageMso=""" & corPai & """ itemSize=""" & TamanhoSubmenu & """>"
    
    ' --- CONTROLE DO TÍTULO: DIRETÓRIO ---
    If TipoCabecalhoSubmenu = ePdrSeparador Then
        xml = xml & "    <menuSeparator id=""sep_dir_sys" & m_idNum & """ title=""DIRETÓRIO""/>"
    Else
        xml = xml & "    <button id=""hdr_dir_sys" & m_idNum & """ label=""DIRETÓRIO"" enabled=""false""/>"
    End If
    
    xml = xml & "    <button id=""btn_sel_sys" & m_idNum & """ label=""Selecionar Novo Arquivo..."" image=""img_procurar"" onAction=""OnAct_SelecionarArquivo"" tag=""" & TagPQ & "|" & CaminhoCompleto & """/>"
    xml = xml & "    <button id=""btn_go_sys" & m_idNum & """ label=""Ir para Arquivo"" image=""img_pastaAberta"" onAction=""OnAct_AbrirPasta"" tag=""" & CaminhoCompleto & """/>"
    xml = xml & "</menu>"
    
    m_SecoesXML(SecaoID) = m_SecoesXML(SecaoID) & xml
    m_idNum = m_idNum + 1
End Sub

' ==========================================================================
' Função:      GerarXML
' Descrição:   Varre todas as seções mapeadas e constrói a string XML final
'              que será injetada no CustomUI do Ribbon.
' Retorno:     String com a formatação XML completa.
' ==========================================================================
Public Function GerarXML() As String
    Dim xmlFinal As String, chaveSecao As Variant
    xmlFinal = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" itemSize=""" & TamanhoMenuPrincipal & """>"
    
    ' Respeita a ordem de criação das seções ao agrupar os nós XML
    For Each chaveSecao In m_SecoesOrdem
        If InStr(1, m_SecoesXML(chaveSecao), "<menu ") > 0 Then xmlFinal = xmlFinal & m_SecoesXML(chaveSecao)
    Next chaveSecao
    
    GerarXML = xmlFinal & "</menu>"
End Function

' =====================================================================
' BLOCO 2: MOTOR DE AÇÕES (Lógica de Cliques e Windows)
' =====================================================================

' ==========================================================================
' Rotina:      ExecutarTogglePadrao
' Descrição:   Ativa ou desativa o uso de um diretório padrão, atualizando 
'              a variável no Power Query.
' Parâmetros:  Tag (String) - Variáveis separadas por pipe ('|')
' ==========================================================================
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

' ==========================================================================
' Rotina:      ExecutarSelecionarPasta
' Descrição:   Abre o FileDialog nativo para seleção de pastas e salva
'              o caminho escolhido no repositório Power Query.
' Parâmetros:  Tag (String) - NomeDaVariavel|CaminhoInicial
' ==========================================================================
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
            ' Assegura que o caminho registrado finalize em barra para operações posteriores de concatenação
            If Right(novaPasta, 1) <> "\" Then novaPasta = novaPasta & "\"
            SalvarVariavelPQ_Interno NOME_CONSULTA_PQ, nomeVariavel, novaPasta
        End If
    End With
End Sub

' ==========================================================================
' Rotina:      ExecutarSelecionarArquivo
' Descrição:   Abre o FileDialog nativo para seleção de arquivos.
' Parâmetros:  Tag (String) - NomeDaVariavel|CaminhoInicial
' ==========================================================================
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
            ' Isola o diretório pai do caminho do arquivo para focar o FileDialog
            pastaInicial = Left(caminhoInicial, InStrRev(caminhoInicial, "\"))
            If Dir(pastaInicial, vbDirectory) <> "" Then .InitialFileName = pastaInicial
        End If
        If .Show = -1 Then
            novoArquivo = .SelectedItems(1)
            SalvarVariavelPQ_Interno NOME_CONSULTA_PQ, nomeVariavel, novoArquivo
        End If
    End With
End Sub

' ==========================================================================
' Rotina:      ExecutarAbrirCaminho
' Descrição:   Integra com o Windows Explorer para focar em pastas ou
'              destacar arquivos específicos.
' Parâmetros:  Tag (String) - Caminho a ser explorado
' ==========================================================================
Public Sub ExecutarAbrirCaminho(Tag As String)
    Dim pastaRaiz As String
    If Tag = "" Then Exit Sub
    On Error Resume Next
    ' Verifica se a string remete a um recurso existente
    If Dir(Tag, vbDirectory) <> "" Then
        ' Avalia os atributos para decidir se abre a pasta ou foca no arquivo
        If (GetAttr(Tag) And vbDirectory) = vbDirectory Then
            Shell "explorer.exe """ & Tag & """", vbNormalFocus
        Else
            Shell "explorer.exe /select,""" & Tag & """", vbNormalFocus
        End If
    Else
        ' Fallback de navegação: se falhou, tenta abrir o diretório pai
        pastaRaiz = Left(Tag, InStrRev(Tag, "\"))
        If Dir(pastaRaiz, vbDirectory) <> "" Then Shell "explorer.exe """ & pastaRaiz & """", vbNormalFocus
    End If
    On Error GoTo 0
End Sub

' =====================================================================
' BLOCO 3: CRUD POWER QUERY (O "Banco de Dados" Embutido e Privado)
' =====================================================================

' ==========================================================================
' Rotina:      SalvarVariavelPQ_Interno
' Descrição:   Camada abstrata que decide se cria ou atualiza uma variável
'              dentro do código M de uma consulta Power Query.
' Parâmetros:  nomeConsulta, nomeVariavel, valor
' ==========================================================================
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

' ==========================================================================
' Função:      LerVariavelPQ_Interno
' Descrição:   Extrai o valor de uma variável específica via Regex dentro
'              do corpo da consulta M.
' Retorno:     Valor em string limpa (sem aspas do código M).
' ==========================================================================
Private Function LerVariavelPQ_Interno(nomeConsulta As String, nomeVariavel As String) As String
    Dim formulaM As String, regex As Object, matches As Object
    LerVariavelPQ_Interno = ""
    On Error Resume Next
    formulaM = ActiveWorkbook.Queries(nomeConsulta).Formula
    On Error GoTo 0
    If formulaM = "" Then Exit Function
    
    ' Padrão de captura do valor mapeado à direita do sinal de igual, ignorando quebras de linha
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\b" & nomeVariavel & "\s*=\s*([^,\r\n]+)"
    regex.IgnoreCase = True
    Set matches = regex.Execute(formulaM)
    
    If matches.Count > 0 Then
        LerVariavelPQ_Interno = Trim(matches(0).SubMatches(0))
        ' Remove a formatação textual estrita embutida pela linguagem M ("")
        If Left(LerVariavelPQ_Interno, 1) = """" Then LerVariavelPQ_Interno = Mid(LerVariavelPQ_Interno, 2, Len(LerVariavelPQ_Interno) - 2)
    End If
End Function

' ==========================================================================
' Rotina:      PrepararConsultaRepositorio
' Descrição:   Garante que a Query base do "banco de dados" exista no arquivo.
' ==========================================================================
Private Sub PrepararConsultaRepositorio(nomeConsulta As String)
    Dim qry As WorkbookQuery
    On Error Resume Next
    Set qry = ActiveWorkbook.Queries(nomeConsulta)
    On Error GoTo 0
    ' Se a query falhar no objeto ou não for encontrada, instanciamos a estrutura base da linguagem M
    If qry Is Nothing Then ActiveWorkbook.Queries.Add Name:=nomeConsulta, Formula:="let" & vbCrLf & "    CriadoEm = """ & Now & """" & vbCrLf & "in" & vbCrLf & "    CriadoEm"
End Sub

' ==========================================================================
' Rotina:      AdicionarVariavelPQ_Interno
' Descrição:   Injeta uma nova linha de variável logo abaixo do bloco "let"
'              utilizando substituição por expressão regular.
' ==========================================================================
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

' ==========================================================================
' Rotina:      AlterarVariavelPQ_Interno
' Descrição:   Substitui de forma cirúrgica apenas o valor de uma variável 
'              existente no código M da consulta.
' ==========================================================================
Private Sub AlterarVariavelPQ_Interno(nomeConsulta As String, nomeVariavel As String, novoValor As String)
    Dim formulaM As String, regex As Object
    formulaM = ActiveWorkbook.Queries(nomeConsulta).Formula
    Set regex = CreateObject("VBScript.RegExp")
    ' Captura a chave em um grupo e o valor em outro para realizar o replace seguro
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

' ==========================================================================
' Função:      ValidarArquivos
' Descrição:   Processa e emite o alerta consolidado caso arquivos da estrutura
'              estejam ausentes ou com extensões/datas divergentes.
' Retorno:     Booleano, True se o fluxo deve prosseguir, False se abortado.
' ==========================================================================
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