# Automação no Active Directory com PowerShell
 Este script em PowerShell fornece uma interface interativa para a administração de usuários no Active Directory (AD). Ele foi desenvolvido para facilitar a gestão de contas de usuários, permitindo que administradores realizem tarefas comuns de forma rápida, segura e padronizada.
A automação do Active Directory é essencial para empresas que desejam reduzir erros humanos, aumentar a produtividade e garantir conformidade nas operações de TI. Este script foi criado para atender a essas necessidades, proporcionando um fluxo de trabalho eficiente para o gerenciamento de contas no AD.

---

## O que este script faz?

- Criação de usuários
- Reset de senha
- Ativação e desativação de contas
- Bloqueio e desbloqueio de usuários

---

## Como funciona?

- O script guia o usuário por um **menu interativo**, permitindo a escolha da ação desejada.
- Para **criação de usuário**, ele valida o nome inserido, corrige a formatação e sugere um login no formato *nome.ultimosobrenome*. Se o login já existir no AD, tenta variações até encontrar um disponível.
- Caso todos os logins possíveis estejam ocupados, o sistema solicita um login manual e verifica sua disponibilidade.
- O script também solicita unidade e departamento, garantindo que o usuário seja criado no local correto com as permissões adequadas. Os departamentos são listados conforme a unidade escolhida, evitando erros de alocação.
- O script também solicita unidade e departamento, garantindo que o usuário seja criado no local correto com as permissões adequadas.
  - O usuário escolhe uma unidade dentre as opções disponíveis.
  - Caso a unidade selecionada seja Corporativo, são listados departamentos específicos dessa unidade.
  - Para as demais unidades (filiais), são exibidos departamentos padronizados, garantindo que cada usuário seja vinculado corretamente.
- No final, exibe um resumo com os dados inseridos e permite ao administrador confirmar ou editar antes da criação do usuário.
- Para reset de senha, o script gera automaticamente uma senha segura de 10 caracteres.
- No caso de bloqueio, desbloqueio, ativação e inativação, o sistema verifica o status atual do usuário antes de realizar qualquer alteração, evitando ações desnecessárias.

---

## Benefícios

- Redução de erros humanos
- Agilidade na administração de contas
- Padronização no gerenciamento de usuários

---

## Por que usar essa automação?

- Este script otimiza a gestão de usuários no Active Directory, reduzindo tempo gasto em tarefas repetitivas e permitindo que a equipe de TI se concentre em atividades estratégicas.

---

## Como usar este script

### Fazendo um Fork do Repositório

1. Acesse o repositório no GitHub.
2. Clique no botão Fork para criar uma cópia do repositório na sua conta.
3. Clone o repositório para sua máquina local:
   ```
   git clone https://github.com/seu-usuario/nome-do-repositorio.git ```
