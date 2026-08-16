Option Explicit

' =====================================================================
' VARIÁVEIS GLOBAIS DE TESTE
' =====================================================================
' Agora temos diretórios padrões independentes!
' Variáveis para armazenar os caminhos padrão, simulando o retorno do Power Query
Public dirPadraoOPC As String
Public dirPadraoOBRIG As String

' Variáveis para armazenar os caminhos específicos de cada arquivo modular mapeado
Public dirAderencia As String
Public dirQualidade As String
Public caminhoQuadro As String

' ==========================================================================
' Rotina:      CriarAmbienteDeTeste
' Descrição:   Gera um ambiente local seguro (Mock) na Área de Trabalho do 
'              usuário, construindo árvores de diretórios e arquivos falsos 
'              necessários para validar as regras de negócio do RibbonMenu.
' Parâmetros:  Nenhum
' ==========================================================================
Public Sub CriarAmbienteDeTeste()
    Dim wsh As Object, fso As Object
    Dim desktopPath As String, pastaTeste As String
    Dim pastaOPC As String, pastaOBRIG As String
    
    ' Instancia o Windows Script Host via Late Binding para localizar dinamicamente a Área de Trabalho do usuário ativo
    Set wsh = CreateObject("WScript.Shell")
    desktopPath = wsh.SpecialFolders("Desktop")
    
    ' Definindo os caminhos para o teste de múltiplos padrões
    pastaTeste = desktopPath & "\Teste_Ribbon_XML\"
    pastaOPC = pastaTeste & "Opcionais\"
    pastaOBRIG = pastaTeste & "Obrigatorios\"
    
    ' Inicializa o FileSystemObject para manipulação de arquivos físicos no SO
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Cria a estrutura de pastas hierárquicas, garantindo que não sobrescreva caso já existam
    If Not fso.FolderExists(pastaTeste) Then fso.CreateFolder pastaTeste
    If Not fso.FolderExists(pastaOPC) Then fso.CreateFolder pastaOPC
    If Not fso.FolderExists(pastaOBRIG) Then fso.CreateFolder pastaOBRIG
    
    ' Cria os arquivos falsos na pasta de Obrigatórios para simular o comportamento das validações de extensão e data
    CriarArquivoFalso fso, pastaOBRIG & "Aderência_Atualizada.xlsx"
    CriarArquivoFalso fso, pastaOBRIG & "Qualidade_Atualizada.xlsx"
    CriarArquivoFalso fso, pastaTeste & "Quadro.xlsb"
    
    ' Simula a leitura das variáveis (Como se viessem do Power Query)
    dirPadraoOPC = pastaOPC
    dirPadraoOBRIG = pastaOBRIG
    
    ' Arquivos apontando para o padrão Obrigatório (O Toggle ficará verde!)
    dirAderencia = pastaOBRIG
    dirQualidade = pastaOBRIG
    caminhoQuadro = pastaTeste & "Quadro.xlsb"
End Sub

' ==========================================================================
' Rotina:      CriarArquivoFalso
' Descrição:   Rotina auxiliar que gera um arquivo de texto genérico (Dummy) 
'              com a extensão solicitada, se ele ainda não existir.
' Parâmetros:  fso (Object) - Instância do FileSystemObject.
'              caminho (String) - Caminho completo onde o arquivo será salvo.
' ==========================================================================
Private Sub CriarArquivoFalso(fso As Object, caminho As String)
    ' Verifica a preexistência para evitar erros de sobrescrita ou arquivos travados
    If Not fso.FileExists(caminho) Then
        Dim f As Object
        ' Cria o arquivo de texto (True permite sobrescrever, embora o FileExists já atue como guarda)
        Set f = fso.CreateTextFile(caminho, True)
        f.WriteLine "Arquivo de teste gerado pelo VBA."
        f.Close
    End If
End Sub