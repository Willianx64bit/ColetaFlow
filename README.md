# ColetaFlow

**Coletor universal de inventário e conferência.**

O ColetaFlow é independente do MobFlow. O coletor local continua disponível sem internet e sem login. A área compartilhada permite que o dono distribua uma planilha, convide usuários e acompanhe a coleta.

## Contas e permissões

Existem somente dois tipos de conta:

- **DONO:** uma única conta proprietária da plataforma. É a única pessoa que cadastra usuários, define senhas iniciais, bloqueia/libera acessos, importa modelos, cria coletas, acompanha participantes e baixa a planilha final.
- **USUÁRIOS:** contas criadas manualmente pelo dono. Entram com e-mail e senha próprios e participam das coletas liberadas.

Não existe gerente, supervisor, cadastro público ou autoinscrição. A conta DONO é protegida pela tabela `cf_platform_owner`; não depende apenas do campo de papel da conta.

## Fluxo compartilhado

1. O dono importa a planilha original e cria uma coleta.
2. O dono confirma a aba e o mapeamento das colunas de código e quantidade.
3. Os usuários entram com suas próprias contas e o código da coleta.
4. Cada dispositivo registra as bipagens localmente no IndexedDB. Se estiver sem internet, mantém a fila pendente.
5. Quando a conexão volta, os eventos são enviados automaticamente.
6. O painel do dono acompanha usuários ativos, finalizados, eventos, códigos e quantidades recebidas.
7. Depois que todos finalizam, o dono encerra a coleta e baixa a planilha consolidada.

O dono também pode reabrir coletas anteriores pelo painel, sem depender apenas do armazenamento local do navegador.

## Integridade e sincronização

- Cada bipagem ou ajuste tem um UUID próprio e é armazenado como evento independente.
- O servidor deduplica pelo UUID. Reenvios idênticos são reconhecidos; reutilizar o mesmo UUID com conteúdo diferente é rejeitado.
- A gravação de lotes é transacional.
- Sincronização, finalização do participante e encerramento da coleta usam bloqueios transacionais para evitar que uma bipagem pendente seja perdida no fechamento.
- A fila offline é persistida separadamente no IndexedDB e as gravações locais são serializadas.
- O coletor local sem login permanece independente da área compartilhada.

## Exportação

- O modelo original é armazenado com a coleta e distribuído aos participantes.
- O mapeamento confirmado pelo dono fica bloqueado depois que um usuário entra, evitando divergência entre dispositivos.
- Arquivos `.xlsx` são regravados com ExcelJS para preservar estilos e estrutura sempre que suportado.
- Arquivos `.xls` e `.csv` usam SheetJS.
- Códigos não encontrados no modelo são contabilizados como não correspondentes; a exportação não altera o arquivo original.

A preservação de recursos Excel avançados (macros, conexões, tabelas dinâmicas e fórmulas especiais) ainda precisa ser validada com modelos reais. Não considerar o produto pronto para operação comercial antes dos testes ponta a ponta.

## Configuração do ambiente de teste

As variáveis abaixo ficam somente no ambiente **Preview** da Vercel, limitadas à branch `feature/shared-collections`:

- `DATABASE_URL`: conexão do branch Neon de teste `br-delicate-cloud-b63jsp4d`.
- `COLETAFLOW_OWNER_EMAIL`: e-mail autorizado da única conta DONO.
- `COLETAFLOW_OWNER_SETUP_KEY`: chave aleatória forte, armazenada como segredo.

O primeiro acesso do dono é feito pelo botão **Primeiro acesso do dono** na tela de login. O endpoint exige a chave, o e-mail autorizado, nome e senha inicial de pelo menos 10 caracteres. A criação é de uso único; depois que existe um proprietário, não é possível criar outro pelo bootstrap.

## Banco de dados

Migrações da branch:

- `001_shared_collections.sql`: tabelas de contas, sessões, coletas, participantes e eventos.
- `002_owner_user_roles.sql`: restringe os papéis a `owner` e `user`.
- `003_independent_scan_events.sql`: mantém cada leitura e ajuste como evento independente.
- `004_template_mapping.sql`: armazena o mapeamento confirmado.
- `005_unique_collection_accounts.sql`: permite nomes iguais e uma participação por conta em cada coleta.

As migrações e testes de banco devem usar exclusivamente o branch Neon de teste, nunca o branch principal.

## Desenvolvimento e segurança

- Repositório: https://github.com/Willianx64bit/ColetaFlow
- Branch de desenvolvimento: `feature/shared-collections`
- Banco Neon de teste: `br-delicate-cloud-b63jsp4d`
- Projeto Vercel: `coletaflow`
- O workflow do GitHub Actions valida sintaxe dos endpoints, round-trip de formatação ExcelJS, parser CSV SheetJS e build TypeScript/Vite.
- O parser SheetJS usa a distribuição oficial 0.20.3, em vez do pacote npm antigo 0.18.5.
- Não alterar `main`, não publicar em produção e não usar o banco principal sem autorização expressa.

**Estado atual:** build e testes automatizados aprovados na branch de desenvolvimento. A validação ponta a ponta com dois dispositivos, operação offline real e planilhas de empresas ainda está pendente.
