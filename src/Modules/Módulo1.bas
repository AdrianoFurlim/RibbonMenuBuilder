Option Explicit

' =====================================================================
' VARIÁVEIS GLOBAIS DE TESTE
' =====================================================================
Public dirPadrao As String
Public dirAderencia As String
Public dirQualidade As String
Public caminhoQuadro As String

' =====================================================================
' ROTINA QUE PREPARA O AMBIENTE (Cria Pastas e Arquivos Fake)
' =====================================================================
Public Sub CriarAmbienteDeTeste()
    Dim wsh As Object, fso As Object
    Dim desktopPath As String, pastaTeste As String
    
    ' Pega o caminho da Área de Trabalho do usuário atual
    Set wsh = CreateObject("WScript.Shell")
    desktopPath = wsh.SpecialFolders("Desktop")
    pastaTeste = desktopPath & "\Teste_Ribbon_XML\"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Cria a pasta raiz se não existir
    If Not fso.FolderExists(pastaTeste) Then
        fso.CreateFolder pastaTeste
    End If
    
    ' Cria os arquivos "falsos" (são arquivos de texto com extensão de Excel)
    ' Como a classe só checa o "Dir" e "FileDateTime", isso funciona perfeitamente e é super rápido!
    CriarArquivoFalso fso, pastaTeste & "Aderência_0108.xlsx"
    CriarArquivoFalso fso, pastaTeste & "Qualidade_0108.xlsx"
    CriarArquivoFalso fso, pastaTeste & "Quadro.xlsb"
    
    ' Alimenta as variáveis globais simulando a leitura do Power Query
    dirPadrao = pastaTeste
    dirAderencia = pastaTeste       ' Igual ao padrão (Botão "Usar Padrão" ficará Verde/Ativado)
    dirQualidade = pastaTeste
    caminhoQuadro = pastaTeste & "Quadro.xlsb"
    
End Sub

' Função de apoio para criar os arquivos físicos no Windows
Private Sub CriarArquivoFalso(fso As Object, caminho As String)
    If Not fso.FileExists(caminho) Then
        Dim f As Object
        Set f = fso.CreateTextFile(caminho, True)
        f.WriteLine "Arquivo de teste gerado pelo VBA para o menu."
        f.Close
    End If
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
