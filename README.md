# Avaliação de Subcontratados
Este projeto foi desenvolvido para facilitar o lançamento e a gestão de pontuações para motoristas e fornecedores, com base em critérios específicos. O sistema permite que o usuário selecione e preencha informações diretamente de uma planilha Excel e faça o lançamento das pontuações em abas distintas para motoristas e fornecedores.


FUNCIONALIDADES PRINCIPAIS

1. Lançamento de Pontuações de Motoristas
O sistema permite a inserção de pontuações para motoristas, com base em critérios específicos. O usuário pode selecionar o motorista, escolher o critério de avaliação, e o sistema calcula automaticamente a pontuação total de acordo com o critério selecionado.

2. Lançamento de Pontuações de Fornecedores
Além de motoristas, o sistema também possibilita o lançamento de pontuações para fornecedores, utilizando os critérios definidos para avaliá-los. O processo de preenchimento é semelhante ao de motoristas, garantindo uma interface consistente.

3. Integração com Planilhas Excel
O sistema lê os dados de motoristas, fornecedores e critérios diretamente de uma planilha Excel. As informações são carregadas automaticamente, e os lançamentos realizados são salvos na mesma planilha, garantindo fácil acesso aos dados.

4. Interface Gráfica Amigável
Desenvolvido em PyQt5, o sistema oferece uma interface gráfica simples e intuitiva, com abas separadas para motoristas e fornecedores, facilitando o uso e a navegação entre diferentes módulos.

5. Tema Escuro
Para melhorar a experiência do usuário, o sistema aplica um tema escuro, que torna a visualização mais agradável durante o uso prolongado.


TECNOLOGIAS UTILIZADAS

1. Python: Linguagem principal utilizada para o desenvolvimento do sistema.

2. PyQt5: Biblioteca utilizada para criar a interface gráfica do aplicativo.

3. Pandas: Utilizado para leitura e manipulação dos dados da planilha Excel.

4. Openpyxl: Utilizado para adicionar novos dados e salvar os lançamentos na planilha Excel.


COMO FUNCIONA

O usuário inicia o sistema e seleciona o arquivo Excel que contém os dados de motoristas, fornecedores e critérios.
Através da interface, o usuário escolhe a aba correspondente (Motoristas ou Fornecedores).
O usuário seleciona o nome do motorista ou fornecedor, escolhe o critério de avaliação e a pontuação é preenchida automaticamente.
O sistema permite que o usuário salve os dados diretamente na planilha Excel, atualizando a aba correspondente com os novos lançamentos.
Uma mensagem de confirmação informa que os dados foram lançados com sucesso.


CONCLUSÃO

Este sistema oferece uma solução eficiente para o lançamento e gerenciamento de pontuações de motoristas e fornecedores. Com a integração de uma interface gráfica amigável e a manipulação direta de arquivos Excel, o processo de avaliação se torna mais prático e organizado.
