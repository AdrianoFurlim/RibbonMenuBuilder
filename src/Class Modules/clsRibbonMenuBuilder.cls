Option Explicit

' =====================================================================
' CONTROLE XML DA RIBBON (Para inserir no Custom UI Editor):
' <dynamicMenu id="tbVerificarArq" label="Verificar Arquivos" image="Definicoes" size="large" getContent="GerarMenuArquivos" invalidateContentOnDrop="true" screentip="Verificar Arquivos" supertip="Abre o menu com a lista de arquivos necessários para verificar pendências."/>
' =====================================================================

' =====================================================================
' VARIÁVEIS INTERNAS (Estado da Classe)
' =====================================================================
' Ícones padronizados (Podem ser alterados via Propriedades)
Private m_IconeOk As String
Private m_IconeAtencao As String
Private m_IconeErro As String
Private m_IconeCinza As String

' Estruturas de armazenamento (Late Binding para evitar dependências)
Private m_SecoesOrdem As Collection      ' Guarda a ordem das seções
Private m_SecoesXML As Object            ' Dicionário (ID da Seção -> XML acumulado)

Private m_ConfigXML As String            ' XML do menu fixo de configuração
Private m_idNum As Integer               ' Contador global para IDs únicos dos botões

' =====================================================================
' INICIALIZAÇÃO DA CLASSE (Construtor)
' =====================================================================
Private Sub Class_Initialize()
    ' Define os ícones padrão
    m_IconeOk = "NotebookColor12"       ' Verde
    m_IconeAtencao = "NotebookColor2"   ' Laranja/Amarelo
    m_IconeErro = "NotebookColor7"      ' Vermelho
    m_IconeCinza = "NotebookColor16"    ' Cinza
    
    ' Inicializa as coleções e dicionários (Late Binding)
    Set m_SecoesOrdem = New Collection
    Set m_SecoesXML = CreateObject("Scripting.Dictionary")
    
    m_idNum = 1
    m_ConfigXML = ""
End Sub

' =====================================================================
' PROPRIEDADES (Permite ao usuário trocar os ícones se desejar)
' =====================================================================
Public Property Let IconeOk(Valor As String): m_IconeOk = Valor: End Property
Public Property Let IconeAtencao(Valor As String): m_IconeAtencao = Valor: End Property
Public Property Let IconeErro(Valor As String): m_IconeErro = Valor: End Property
Public Property Let IconeCinza(Valor As String): m_IconeCinza = Valor: End Property

' =====================================================================
' 1. MÉTODO: ADICIONAR MENU DE CONFIGURAÇÃO (Topo)
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

' =====================================================================
' 2. MÉTODO: CRIAR SEÇÃO (Caixas organizadoras)
' =====================================================================
Public Sub CriarSecao(ID As String, Titulo As String, Optional IncluirSeparador As Boolean = True)
    Dim xmlCabecalho As String
    
    ' Se a seção já existe, aborta para não duplicar
    If m_SecoesXML.exists(ID) Then Exit Sub
    
    ' Inicia o XML da seção com o separador (se pedido) e o botão de título inativo
    If IncluirSeparador Then xmlCabecalho = "<menuSeparator id=""sep_" & ID & """/>"
    xmlCabecalho = xmlCabecalho & "<button id=""hdr_" & ID & """ label=""" & Titulo & """ enabled=""false""/>"
    
    ' Registra a seção
    m_SecoesOrdem.Add ID
    m_SecoesXML.Add ID, xmlCabecalho
End Sub

' =====================================================================
' 3. MÉTODO: ADICIONAR ARQUIVO MODULAR (Pastas)
' =====================================================================
Public Sub AdicionarArquivoModular(SecaoID As String, Nome As String, Prefixo As String, Extensao As String, DiretorioAlvo As String, Obrigatorio As Boolean, TagPQ As String, UsaPadrao As Boolean)
    Dim arquivoAtual As String, arquivoMaisRecente As String
    Dim dataAtual As Date, dataModificacao As Date
    Dim extReal As String, xml As String, caminho As String
    Dim corPai As String, lblPai As String, tagGo As String
    Dim lblUsarPadrao As String, btnSelecionarHabilitado As String, iconePadrao As String
    
    ' Se a seção não foi criada antes, aborta
    If Not m_SecoesXML.exists(SecaoID) Then Exit Sub
    
    ' --- Lógica do Toggle (Interruptor visual) ---
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
    
    ' Motor de Busca no Windows
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
    
    ' --- ESTADO 1: ARQUIVO NÃO ENCONTRADO ---
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
        
    ' --- ESTADO 2: ARQUIVO ENCONTRADO ---
    Else
        extReal = LCase(Mid(arquivoMaisRecente, InStrRev(arquivoMaisRecente, ".")))
        lblPai = Nome
        tagGo = caminho & arquivoMaisRecente
        
        Dim corExt As String, textoExt As String
        corExt = IIf(extReal = LCase(Extensao), m_IconeOk, m_IconeErro)
        textoExt = "Extensão: " & extReal & IIf(extReal = LCase(Extensao), "", " (Incorreta)")
        
        Dim corData As String, textoData As String
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
    
    ' Adiciona o XML gerado à caixinha da Seção correspondente
    m_SecoesXML(SecaoID) = m_SecoesXML(SecaoID) & xml
    m_idNum = m_idNum + 1
End Sub

' =====================================================================
' 4. MÉTODO: ADICIONAR ARQUIVO DIRETO (Sistemas)
' =====================================================================
Public Sub AdicionarArquivoDireto(SecaoID As String, Nome As String, CaminhoCompleto As String, Obrigatorio As Boolean, TagPQ As String)
    Dim xml As String, corPai As String, lblPai As String
    
    ' Se a seção não foi criada antes, aborta
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
    
    ' Adiciona o XML gerado à caixinha da Seção correspondente
    m_SecoesXML(SecaoID) = m_SecoesXML(SecaoID) & xml
    m_idNum = m_idNum + 1
End Sub

' =====================================================================
' 5. MÉTODO: GERAR XML FINAL
' =====================================================================
Public Function GerarXML() As String
    Dim xmlFinal As String
    Dim chaveSecao As Variant
    
    xmlFinal = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" itemSize=""large"">"
    
    ' Adiciona a Configuração (se foi criada)
    If m_ConfigXML <> "" Then
        xmlFinal = xmlFinal & m_ConfigXML
    End If
    
    ' Itera pela coleção de ordens para garantir que o XML saia na mesma ordem em que as seções foram criadas
    For Each chaveSecao In m_SecoesOrdem
        ' Só adiciona a seção se ela contiver algo além do próprio título (para não mostrar seções vazias)
        ' O InStr verifica se tem algum "<menu " depois do cabeçalho
        If InStr(1, m_SecoesXML(chaveSecao), "<menu ") > 0 Then
            xmlFinal = xmlFinal & m_SecoesXML(chaveSecao)
        End If
    Next chaveSecao
    
    xmlFinal = xmlFinal & "</menu>"
    GerarXML = xmlFinal
End Function