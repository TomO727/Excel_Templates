📊 Modelos de Planilhas Excel e Códigos VBA 📊

Bem-vindo ao repositório de modelos e códigos VBA para Excel! 
Aqui você encontrará diversos templates de planilhas e módulos VBA prontos para uso, ideais tanto para uso direto quanto para servir de exemplo ou inspiração para outros projetos. 
Este repositório é mantido com o intuito de compartilhar recursos úteis e facilitar o trabalho de quem utiliza Excel e VBA para automação e análise de dados.

Sobre o Repositório 📝
Este repositório foi criado com a ideia de reunir e compartilhar:

Templates de Planilhas Excel: 
Planilhas organizadas e prontas para o uso em diversas situações, que podem ser adaptadas para diferentes cenários e necessidades.

Módulos e Funções VBA: 
Código VBA reutilizável, com functions, subrotinas, matrizes, classes e mais. Ideal para quem deseja aprender VBA ou otimizar processos no Excel.

Criador: Antonio Oliveira Silva
Usuário: Tony O. (TomO727)

Estrutura do Repositório 📂
O repositório está organizado em pastas e arquivos para facilitar a navegação e uso:

/Templates: Contém templates prontos de planilhas Excel, organizados por categoria. Alguns exemplos incluem:

Financeiro: 
Modelos para orçamento, controle de caixa, fluxo de caixa.

Gestão de Projetos: 
Planilhas para cronogramas, acompanhamento de tarefas e status de projetos.
Relatórios e Análises: Planilhas prontas para gráficos e dashboards.
/VBA_Modules: 
Módulos com funções e rotinas VBA que podem ser importadas em projetos existentes. Exemplos:

Funções de Manipulação de Dados: 
Funções para formatar dados, calcular valores ou executar filtros.
Matrizes e Classes: 
Módulos para trabalhar com estruturas de dados avançadas e classes VBA.
Automação de Tarefas: 
Rotinas para automatizar atividades repetitivas, como geração de relatórios e preenchimento de dados.

Como Utilizar 🚀
Usando Templates de Planilhas
Navegue até a pasta /Templates e escolha o template que melhor se adapta à sua necessidade.
Baixe o arquivo .xlsx ou .xlsm e abra-o no Excel.
Adapte o conteúdo conforme necessário e comece a usar.

Importando Módulos VBA
No Excel, pressione ALT + F11 para abrir o editor VBA.
No menu, clique em Arquivo > Importar arquivo e selecione o módulo VBA que você deseja importar.
Após a importação, o módulo estará disponível para uso no projeto atual.

Exemplos de Código 📋
Abaixo estão alguns exemplos de funções e rotinas VBA disponíveis no repositório:

Função para Somar Valores com Condições
vba
Copiar código
Function SomarComCondicao(rng As Range, condicao As String) As Double
    Dim cel As Range
    Dim soma As Double
    soma = 0
    For Each cel In rng
        If cel.Value = condicao Then
            soma = soma + cel.Offset(0, 1).Value
        End If
    Next cel
    SomarComCondicao = soma
End Function

Macro para Gerar Relatório de Vendas
vba
Copiar código
Sub GerarRelatorioVendas()
    ' Configura a planilha e cria um novo relatório de vendas
    Sheets.Add.Name = "Relatório de Vendas"
    With Sheets("Relatório de Vendas")
        .Range("A1").Value = "Data"
        .Range("B1").Value = "Produto"
        .Range("C1").Value = "Quantidade"
    End With
    ' Código para preencher as células com dados simulados
End Sub
Esses são apenas alguns exemplos! Explore o repositório para mais funções e macros que podem ajudar nos seus projetos.

Contribuições 💡
Se você tiver algum código VBA, planilha ou sugestão de melhoria que queira compartilhar, fique à vontade para contribuir! 
Abra um pull request com o seu código, ou, se preferir, entre em contato para sugestões e melhorias.
