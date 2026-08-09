Option Explicit

' =====================================================================
' CALLBACKS DA RIBBON (Ponte de comunicação com a Classe)
' =====================================================================

' Alterna entre usar o Padrão ou não
Public Sub OnAct_UsarPadrao(control As IRibbonControl)
    Dim menu As New clsRibbonMenuBuilder
    menu.ExecutarTogglePadrao control.Tag
End Sub

' Abre o Seletor de Pastas
Public Sub OnAct_SelecionarPasta(control As IRibbonControl)
    Dim menu As New clsRibbonMenuBuilder
    menu.ExecutarSelecionarPasta control.Tag
End Sub

' Abre o Seletor de Arquivos (Para os de Sistema)
Public Sub OnAct_SelecionarArquivo(control As IRibbonControl)
    Dim menu As New clsRibbonMenuBuilder
    menu.ExecutarSelecionarArquivo control.Tag
End Sub

' Abre a pasta física no Windows (destacando o arquivo se existir)
Public Sub OnAct_AbrirPasta(control As IRibbonControl)
    Dim menu As New clsRibbonMenuBuilder
    menu.ExecutarAbrirCaminho control.Tag
End Sub

' Função amortecedora para botões sem ação (ex: ícone de Data e Extensão)
Public Sub MacroVazia(control As IRibbonControl)
    ' Faz nada
End Sub


' =====================================================================
' GERAÇÃO DO MENU (Chamada pela Ribbon)
' =====================================================================
Public Function GerarMenuArquivos(control As IRibbonControl, ByRef returnedVal)
    ' 0. Garante que as variáveis e arquivos existem antes de desenhar o menu
    If dirPadrao = "" Then Call CriarAmbienteDeTeste

    Dim construtor As New clsRibbonMenuBuilder
    
    ' 1. Cria o bloco fixo de configurações
    construtor.AdicionarMenuConfiguracao DirAtual:=dirPadrao, NomeVariavelPQ:="dirPadrão"
    
    ' 2. Cria as caixinhas (Seções) na ordem em que deseja exibir
    construtor.CriarSecao ID:="OPC", Titulo:="ARQUIVOS OPCIONAIS", IncluirSeparador:=True
    construtor.CriarSecao ID:="OBRIG", Titulo:="ARQUIVOS OBRIGATÓRIOS", IncluirSeparador:=True
    construtor.CriarSecao ID:="SYS", Titulo:="ARQUIVOS DE SISTEMA", IncluirSeparador:=True
    
    ' 3. Adiciona os arquivos
    construtor.AdicionarArquivoModular _
        SecaoID:="OBRIG", _
        Nome:="Aderência (DTO)", _
        Prefixo:="Aderência", _
        Extensao:=".xlsx", _
        DiretorioAlvo:=dirAderencia, _
        Obrigatorio:=True, _
        TagPQ:="dirAderencia", _
        UsaPadrao:=(dirAderencia = dirPadrao)

    construtor.AdicionarArquivoModular _
        SecaoID:="OBRIG", _
        Nome:="Qualidade (DTO)", _
        Prefixo:="Qualidade", _
        Extensao:=".xlsx", _
        DiretorioAlvo:=dirQualidade, _
        Obrigatorio:=True, _
        TagPQ:="dirQualidade", _
        UsaPadrao:=(dirQualidade = dirPadrao)

    construtor.AdicionarArquivoDireto _
        SecaoID:="SYS", _
        Nome:="Quadro CD", _
        CaminhoCompleto:=caminhoQuadro, _
        Obrigatorio:=True, _
        TagPQ:="dirQuadro"
        
    ' 4. Retorna o XML processado para a Ribbon
    returnedVal = construtor.GerarXML()
End Function
