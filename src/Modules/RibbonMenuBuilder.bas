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
' Public Function GerarMenuArquivos(control As IRibbonControl, ByRef returnedVal)
    ' ' 0. Garante que as variáveis e arquivos existem antes de desenhar o menu
    ' If dirPadraoOBRIG = "" Then Call CriarAmbienteDeTeste
'
    ' Dim construtor As New clsRibbonMenuBuilder
    '
    ' ' Tamanho base do menu
    ' construtor.TamanhoMenuPrincipal = "large"
    '
    ' ' 1. Cria as caixinhas (Seções) na ordem em que deseja exibir
    ' construtor.CriarSecao ID:="OPC", Titulo:="ARQUIVOS OPCIONAIS", TipoCabecalho:=ePdrSeparador
    ' construtor.CriarSecao ID:="OBRIG", Titulo:="ARQUIVOS OBRIGATÓRIOS", TipoCabecalho:=ePdrSeparador
    ' construtor.CriarSecao ID:="SYS", Titulo:="ARQUIVOS DE SISTEMA", TipoCabecalho:=ePdrSeparador
    '
    ' ' 2. Adiciona os Diretórios Padrões dentro das suas respectivas seções
    ' construtor.AdicionarDiretorioPadrao _
        ' SecaoID:="OPC", _
        ' Titulo:="Diretório Padrão (Opcionais)", _
        ' VarPQ:="dirPadraoOPC", _
        ' DirAtual:=dirPadraoOPC
        '
    ' construtor.AdicionarDiretorioPadrao _
        ' SecaoID:="OBRIG", _
        ' Titulo:="Diretório Padrão", _
        ' VarPQ:="dirPadraoOBRIG", _
        ' DirAtual:=dirPadraoOBRIG, _
        ' TamanhoSubmenu:="normal"
    '
    ' ' 3. Adiciona os arquivos modulares com os NOVOS parâmetros
    ' construtor.AdicionarArquivoModular _
        ' SecaoID:="OBRIG", _
        ' Nome:="Aderência (DTO)", _
        ' NomeBusca:="Aderência", _
        ' Extensao:=".xlsx", _
        ' DiretorioAlvo:=dirAderencia, _
        ' TagVarEspecifica:="dirAderencia", _
        ' TagVarPadrao:="dirPadraoOBRIG", _
        ' ValorPadrao:=dirPadraoOBRIG, _
        ' AntigoEm:=v1Dia, _
        ' TamanhoSubmenu:="normal"
'
    ' construtor.AdicionarArquivoModular _
        ' SecaoID:="OBRIG", _
        ' Nome:="Qualidade (DTO)", _
        ' NomeBusca:="Qualidade", _
        ' Extensao:=".xlsx", _
        ' DiretorioAlvo:=dirQualidade, _
        ' TagVarEspecifica:="dirQualidade", _
        ' TagVarPadrao:="dirPadraoOBRIG", _
        ' ValorPadrao:=dirPadraoOBRIG, _
        ' AntigoEm:=v1Dia, _
        ' TamanhoSubmenu:="normal"
'
    ' ' 4. Adiciona arquivos de sistema
    ' construtor.AdicionarArquivoDireto _
        ' SecaoID:="SYS", _
        ' Nome:="Quadro CD", _
        ' CaminhoCompleto:=caminhoQuadro, _
        ' TagPQ:="dirQuadro", _
        ' TamanhoSubmenu:="normal"
        '
    ' ' 5. Retorna o XML processado para a Ribbon
    ' returnedVal = construtor.GerarXML()
' End Function


Public Function GerarMenuArquivos(control As IRibbonControl, ByRef returnedVal)
    Dim menu As clsRibbonMenuBuilder
    Set menu = ConfigurarMenuArquivos() ' Puxa a classe já carregada
    
    returnedVal = menu.GerarXML()       ' Cospe o XML para a tela
End Function


Public Function ConfigurarMenuArquivos() As clsRibbonMenuBuilder
    ' 0. Garante que as variáveis e arquivos existem antes de desenhar o menu
    If dirPadraoOBRIG = "" Then Call CriarAmbienteDeTeste

    Dim construtor As New clsRibbonMenuBuilder
    
    ' Tamanho base do menu
    construtor.TamanhoMenuPrincipal = "large"
    
    ' 1. Cria as caixinhas (Seções) na ordem em que deseja exibir
    construtor.CriarSecao ID:="OPC", Titulo:="ARQUIVOS OPCIONAIS", TipoCabecalho:=ePdrSeparador
    construtor.CriarSecao ID:="OBRIG", Titulo:="ARQUIVOS OBRIGATÓRIOS", TipoCabecalho:=ePdrSeparador
    construtor.CriarSecao ID:="SYS", Titulo:="ARQUIVOS DE SISTEMA", TipoCabecalho:=ePdrSeparador
    
    ' 2. Adiciona os Diretórios Padrões dentro das suas respectivas seções
    construtor.AdicionarDiretorioPadrao _
        SecaoID:="OPC", _
        Titulo:="Diretório Padrão (Opcionais)", _
        VarPQ:="dirPadraoOPC", _
        DirAtual:=dirPadraoOPC
        
    construtor.AdicionarDiretorioPadrao _
        SecaoID:="OBRIG", _
        Titulo:="Diretório Padrão", _
        VarPQ:="dirPadraoOBRIG", _
        DirAtual:=dirPadraoOBRIG, _
        TamanhoSubmenu:="normal"
    
    ' 3. Adiciona os arquivos modulares com os NOVOS parâmetros
    construtor.AdicionarArquivoModular _
        SecaoID:="OBRIG", _
        Nome:="Aderência (DTO)", _
        NomeBusca:="Aderência", _
        Extensao:=".xlsx", _
        DiretorioAlvo:=dirAderencia, _
        TagVarEspecifica:="dirAderencia", _
        TagVarPadrao:="dirPadraoOBRIG", _
        ValorPadrao:=dirPadraoOBRIG, _
        AntigoEm:=v10Min, _
        TamanhoSubmenu:="normal", _
        TipoCabecalhoSubmenu:=eBtnSeparador

    construtor.AdicionarArquivoModular _
        SecaoID:="OBRIG", _
        Nome:="Qualidade (DTO)", _
        NomeBusca:="Qualidade", _
        Extensao:=".xlsx", _
        DiretorioAlvo:=dirQualidade, _
        TagVarEspecifica:="dirQualidade", _
        TagVarPadrao:="dirPadraoOBRIG", _
        ValorPadrao:=dirPadraoOBRIG, _
        AntigoEm:=v1Dia, _
        TamanhoSubmenu:="normal"

    ' 4. Adiciona arquivos de sistema
    construtor.AdicionarArquivoDireto _
        SecaoID:="SYS", _
        Nome:="Quadro CD", _
        CaminhoCompleto:=caminhoQuadro, _
        TagPQ:="dirQuadro", _
        TamanhoSubmenu:="normal"
    
    ' No fim, retorna a classe carregada e mastigada
    Set ConfigurarMenuArquivos = construtor
End Function


Sub TesteValidacao() '(Optional ByRef V As Boolean)
    Dim menu As clsRibbonMenuBuilder
    Set menu = ConfigurarMenuArquivos() ' Puxa a MESMA classe já carregada
    
    ' Chama a validação nativa da classe
    If Not menu.ValidarArquivos() Then
        Exit Sub ' Aborta se o usuário clicou em NÃO
    End If
    
    ' ... continua com a sua rotina de AtualizarTudo (cronômetro, RefreshAll, etc) ...
End Sub






