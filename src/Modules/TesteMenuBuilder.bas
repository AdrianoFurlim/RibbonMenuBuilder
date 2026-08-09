Option Explicit

' =====================================================================
' VARIÁVEIS GLOBAIS DE TESTE
' =====================================================================
' Agora temos diretórios padrões independentes!
Public dirPadraoOPC As String
Public dirPadraoOBRIG As String

Public dirAderencia As String
Public dirQualidade As String
Public caminhoQuadro As String

' =====================================================================
' ROTINA QUE PREPARA O AMBIENTE (Cria Pastas e Arquivos Fake)
' =====================================================================
Public Sub CriarAmbienteDeTeste()
    Dim wsh As Object, fso As Object
    Dim desktopPath As String, pastaTeste As String
    Dim pastaOPC As String, pastaOBRIG As String
    
    Set wsh = CreateObject("WScript.Shell")
    desktopPath = wsh.SpecialFolders("Desktop")
    
    ' Definindo os caminhos para o teste de múltiplos padrões
    pastaTeste = desktopPath & "\Teste_Ribbon_XML\"
    pastaOPC = pastaTeste & "Opcionais\"
    pastaOBRIG = pastaTeste & "Obrigatorios\"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Cria a estrutura de pastas
    If Not fso.FolderExists(pastaTeste) Then fso.CreateFolder pastaTeste
    If Not fso.FolderExists(pastaOPC) Then fso.CreateFolder pastaOPC
    If Not fso.FolderExists(pastaOBRIG) Then fso.CreateFolder pastaOBRIG
    
    ' Cria os arquivos falsos na pasta de Obrigatórios
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

Private Sub CriarArquivoFalso(fso As Object, caminho As String)
    If Not fso.FileExists(caminho) Then
        Dim f As Object
        Set f = fso.CreateTextFile(caminho, True)
        f.WriteLine "Arquivo de teste gerado pelo VBA."
        f.Close
    End If
End Sub


