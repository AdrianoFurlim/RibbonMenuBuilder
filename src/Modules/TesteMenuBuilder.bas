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


