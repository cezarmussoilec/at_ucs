# Automatização Universidade Corporativa Senior (UCS)

Código Python Selenium para automatização de criação de usuário e atribuição em cursos na plataforma UCS

Antes de executar:
- Necessário preencher o arquivo credenciais_login.txt com o usuário com perfil de supervisor na plataforma, que irá permitir a criação dos usuários e atribuições nos cursos.
- Necessário preencher o arquivo senha_padrao.txt com a senha que será utilizada para a criação dos novos usuários.

Como funciona?
- Ao executar a aplicação é aberta uma janela solicitando o caminho da planilha que contém as informações dos usuários. Ao selecionar a planilha a automatização já se inicia.
- A aplicação usa como base para verificar os usuário uma coluna "Status" na planilha, se o status for "C" o usuário é ignorado, caso contrário será verificado se o usuário já existe e se está atribuídos nos devidos cursos que foram solicitados.
- Se o usuário já existir e estiver atribuído em todos os cursos solicitados, define o status como "C" e segue para o próximo usuário.
- Se o usuário já existir mas não estiver em algum ou nenhum curso escolhido, ele será atribuído no(s) curso(s) em questão, define o status como "C" e seguirá para o próximo usuário.
- Se o usuário não existir, ele será criado e atribuídos nos cursos escolhidos, define o status como "C" e segue para o próximo usuário.

Atualizações com melhorias virão.
