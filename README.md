# RibbonMenuBuilder
![Visão Geral do Projeto](docs/imgExemplo1.png)

![VBA](https://img.shields.io/badge/Language-VBA-green)
![Excel](https://img.shields.io/badge/Platform-Excel-blue)
![License](https://img.shields.io/badge/License-MIT-blue)

Um framework VBA Orientado a Objetos (OOP) para o Microsoft Excel, desenvolvido para gerar menus dinâmicos na Ribbon (Faixa de Opções). O projeto foca no gerenciamento visual de arquivos, validação de integridade de bases de dados e integração nativa com variáveis do Power Query, eliminando a necessidade de planilhas de configuração externas.

## Sobre o Projeto

O **RibbonMenuBuilder** resolve a complexidade de gerenciar múltiplos arquivos de origem (Excel, CSV, XLSB) que alimentam consultas no Power Query. Em vez de menus estáticos, o sistema gera uma interface dinâmica em tempo real (via XML) que informa o status de cada arquivo usando indicadores visuais (Sinalização Verde, Amarelo e Vermelho) baseados em regras de negócio como: extensão correta, existência no diretório e tempo desde a última modificação.

## Principais Funcionalidades

* **Renderização Dinâmica:** Criação programática de Seções, Diretórios Padrões e Itens de Arquivo.
* **Validação de Integridade (Motor Windows):** Verifica automaticamente existência, extensão e "idade" do arquivo (baseado em um sistema de Enumeração de Tempo: de 10 minutos a 1 semana).
* **CRUD Power Query Integrado:** Lê, salva e altera caminhos de diretórios diretamente no código M das consultas (`let...in`), utilizando RegEx para manipulação cirúrgica.
* **Sistema de "Usar Padrão":** Toggle inteligente para alternar entre diretórios específicos e diretórios mestres.
* **Trava de Segurança:** Método `.ValidarArquivos()` nativo para bloquear execuções críticas caso bases estejam faltantes ou obsoletas.

## Arquitetura

O sistema utiliza o padrão de projeto *Builder*, encapsulando a complexidade em uma classe mestra (`clsRibbonMenuBuilder`):
1. **Casca (XML):** Gatilhos dinâmicos via `invalidateContentOnDrop`.
2. **Cérebro (Classe):** Lógica de estado, IO do Windows, geração de XML e motor RegEx.
3. **Ponte (Callbacks):** Módulos padrão que roteiam as ações da interface para a classe.

## Pré-requisitos

* Microsoft Excel com macros habilitadas (`.xlsm` ou `.xlsb`).
* Referências habilitadas no VBA:
    * `Microsoft Office 16.0 Object Library` (ou versão instalada).
    * `Microsoft Scripting Runtime`.

## Como Usar

### 1. Desenhando o Menu (VBA)

Instancie a classe no módulo responsável pelo callback da Ribbon:

~~~vba
Public Function ConfigurarMenuArquivos() As clsRibbonMenuBuilder
    Dim construtor As New clsRibbonMenuBuilder
    
    ' Define o tamanho do menu principal
    construtor.TamanhoMenuPrincipal = "large"
    
    ' Cria Seção
    construtor.CriarSecao ID:="OBRIG", Titulo:="ARQUIVOS OBRIGATÓRIOS", TipoCabecalho:=ePdrSeparador
    
    ' Adiciona Arquivo Modular
    construtor.AdicionarArquivoModular _
        SecaoID:="OBRIG", _
        Nome:="Aderência (DTO)", _
        NomeBusca:="Aderência", _
        Extensao:=".xlsx", _
        DiretorioAlvo:="C:\Pasta\Dados\", _
        TagVarEspecifica:="dirAderencia", _
        TagVarPadrao:="dirPadraoOBRIG", _
        ValorPadrao:="C:\Pasta\Padrao\", _
        AntigoEm:=v1Dia

    Set ConfigurarMenuArquivos = construtor
End Function
~~~

### 2. Validação antes de executar macros

~~~vba
Sub AtualizarBases() 
    Dim menu As clsRibbonMenuBuilder
    Set menu = ConfigurarMenuArquivos() 
    
    ' Bloqueia execução caso arquivos estejam fora do padrão
    If Not menu.ValidarArquivos() Then Exit Sub 
    
    ' Prossegue com o processamento
    ActiveWorkbook.RefreshAll
End Sub
~~~

## Referência da API

| Método | Descrição | Parâmetros |
| :--- | :--- | :--- |
| `CriarSecao` | Cria uma nova seção no menu | ID, Titulo, TipoCabecalho |
| `AdicionarArquivoModular` | Valida e cria botão de arquivo | SecaoID, Nome, NomeBusca, Extensao, AntigoEm |
| `GerarXML` | Compila o XML final para a Ribbon | - |
| `ValidarArquivos` | Valida integridade e alerta o usuário | - |

## Licença

Este projeto está sob a licença [MIT](https://choosealicense.com/licenses/mit/). 

**MIT License**

Copyright (c) 2026 [Adriano Furtado Lima]

A permissão é concedida, a título gratuito, a qualquer pessoa que obtenha uma cópia deste software e dos arquivos de documentação associados (o "Software"), para lidar com o Software sem restrições, incluindo, sem limitação, os direitos de uso, cópia, modificação, fusão, publicação, distribuição, sublicenciamento e/ou venda de cópias do Software, mediante as seguintes condições:

O aviso de copyright acima e este aviso de permissão devem ser incluídos em todas as cópias ou partes substanciais do Software.

O SOFTWARE É FORNECIDO "COMO ESTÁ", SEM GARANTIA DE QUALQUER TIPO, EXPRESSA OU IMPLÍCITA, INCLUINDO MAS NÃO SE LIMITANDO A GARANTIAS DE COMERCIALIZAÇÃO, ADEQUAÇÃO A UM FIM ESPECÍFICO E NÃO INFRAÇÃO. EM NENHUM CASO OS AUTORES OU DETENTORES DE DIREITOS AUTORAIS SERÃO RESPONSÁVEIS POR QUALQUER RECLAMAÇÃO, DANOS OU OUTRA RESPONSABILIDADE, SEJA EM UMA AÇÃO DE CONTRATO, DELITO OU DE OUTRA FORMA, DECORRENTE DE, OU EM CONEXÃO COM O SOFTWARE OU O USO OU OUTRAS NEGOCIAÇÕES NO SOFTWARE.