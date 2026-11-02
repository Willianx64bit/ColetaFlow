# ColetaFlow

**Coletor inteligente universal de produtos.**

## Desenvolvimento compartilhado (branch feature/shared-collections)

O projeto está evoluindo para dois tipos de conta, sem cadastro público:

- **Dono:** existe um único proprietário da plataforma, identificado por `cf_platform_owner`. Somente ele cria coletas, cadastra usuários e bloqueia/libera acessos.
- **Usuários:** contas criadas manualmente pelo dono, com e-mail e senha inicial próprios.
- Não existe conta de gerente ou supervisor. O endpoint legado de cadastro público permanece bloqueado.
- O modo local continua independente: não exige login nem internet para coletar e exportar neste dispositivo.
- Nas coletas compartilhadas, as leituras e ajustes são eventos independentes com UUID. O cliente mantém eventos pendentes no IndexedDB e tenta sincronizá-los quando a conexão volta. O servidor usa o UUID para evitar duplicação e rejeita reutilização do mesmo identificador com conteúdo diferente.
- O painel apresenta o progresso por participante e a situação de conclusão.

### Configuração inicial do proprietário (ambiente Preview)

Configure estas variáveis somente no ambiente **Preview** da Vercel durante os testes:

- `DATABASE_URL`: conexão do branch Neon de teste `br-delicate-cloud-b63jsp4d`.
- `COLETAFLOW_OWNER_EMAIL`: e-mail autorizado para a única conta de dono.
- `COLETAFLOW_OWNER_SETUP_KEY`: chave aleatória forte, criada pelo proprietário e mantida como segredo.

Depois de configurar as variáveis, abra a tela de login, escolha **Primeiro acesso do dono** e informe a chave, o e-mail autorizado, o nome e uma senha inicial com pelo menos 10 caracteres. A configuração é de uso único; depois que a conta do dono existir, o endpoint não cria outro proprietário.

### Migrações

- `001_shared_collections.sql`: estrutura base completa para uma instalação nova.
- `002_owner_user_roles.sql`: normaliza os papéis para `owner` e `user`.
- `003_independent_scan_events.sql`: habilita eventos de ajuste assinados sem apagar eventos anteriores.
- `004_template_mapping.sql`: salva o mapeamento confirmado pelo dono para a exportação consolidada.
- `005_unique_collection_accounts.sql`: permite nomes iguais e reentrada segura usando o ID da conta.

Aplicações de migração e testes de banco devem usar o branch Neon de teste, nunca o branch principal.

**Implementado nesta branch, ainda aguardando validação ponta a ponta:** envio e distribuição do modelo original, confirmação do mapeamento pelo dono e exportação consolidada usando os eventos recebidos. A fidelidade de formatação de arquivos Excel complexos e o comportamento offline em vários dispositivos ainda precisam de testes reais. Não usar para inventário comercial até esses testes serem concluídos.


## Estado do desenvolvimento

**MVP com coleta sem cadastro prévio: importa um modelo vazio, recebe códigos e quantidades e exporta uma cópia preenchida.**

Já existe uma primeira interface React + TypeScript com:
- Importação de modelos vazios ou com cabeçalho em XLSX, XLS e CSV.
- Sugestão automática e mapeamento manual somente das colunas de código e quantidade.
- Entrada por leitor físico em modo teclado (HID).
- Leitura pela câmera usando ZXing.
- Coleta sem cadastro prévio: não importa catálogo de produtos, não procura correspondências e não exige nomes, descrições ou saldo esperado. O operador lê um código e informa a quantidade.
- Agrupamento das quantidades por código lido; códigos repetidos somam à quantidade já registrada.
- Ajuste manual das quantidades e exportação para XLSX.
- Exportação de uma linha por código lido, preenchendo somente as colunas mapeadas de código e quantidade no modelo.
- Contagem guardada localmente no navegador e recuperável ao reimportar o mesmo arquivo.
- Workflow do GitHub Actions para verificar instalação e build.

**Limitação atual:** ainda é uma primeira base de desenvolvimento, não uma versão validada para operação comercial. É necessário testar modelos reais com diferentes cabeçalhos, posições de coluna, abas, fórmulas e formatações. A leitura da câmera exige HTTPS e permissão do dispositivo. A sessão guarda as contagens localmente, mas o arquivo de origem precisa ser reimportado após recarregar a página. Preservação avançada de macros, fórmulas, estilos especiais e layouts complexos ainda precisa de testes.

Próximos passos: validar o build, testar com planilhas reais variadas, reforçar a persistência offline e implementar instalação PWA/Android.

O ColetaFlow é um aplicativo independente para importar um modelo de planilha vazio, coletar códigos de barras e quantidades sem cadastro prévio e devolver uma cópia preenchida no formato esperado pela empresa.

> **Princípio do produto:** a empresa mantém seu próprio modelo; o ColetaFlow adapta a coleta e a exportação a esse modelo, sem exigir uma planilha padrão única.

## Objetivo

Permitir que uma pessoa ou equipe importe um modelo de planilha vazio, leia códigos de barras, informe as quantidades encontradas e exporte uma cópia preenchida no formato esperado pela empresa.

Fluxo principal:

**Importar modelo vazio → mapear coluna de código e coluna de quantidade → bipar código → informar quantidade → conferir códigos e quantidades → exportar uma cópia preenchida.**

## Funcionalidades planejadas

### 1. Importação de arquivos
- Primeira versão: arquivos Excel `.xlsx` e CSV.
- Identificar abas, cabeçalhos e linhas de produtos.
- Reconhecer as colunas de código de barras/código e quantidade.
- Mostrar uma prévia antes de iniciar.
- Não alterar o arquivo original durante a importação.

### 2. Mapeamento universal de colunas
- Tentar reconhecer automaticamente os cabeçalhos, inclusive variações como EAN, GTIN, código, barras, referência, produto, descrição, estoque e quantidade.
- Quando houver dúvida, pedir ao usuário que selecione a coluna correta.
- Permitir definir a coluna de destino da quantidade e, quando aplicável, a regra de atualização.
- Salvar um perfil de mapeamento reutilizável por modelo/empresa, sempre permitindo revisão.
- Nunca adivinhar silenciosamente uma coluna quando isso puder causar escrita incorreta.

### 3. Coletor e leitura
- Leitura pela câmera do celular.
- Suporte a leitores físicos Bluetooth/USB que funcionem como teclado (HID), quando disponíveis no dispositivo.
- Operação contínua, com foco pronto para a próxima leitura.
- Quantidade grande e legível na tela, feedback visual e sonoro configurável.
- Após a leitura, o operador informa a quantidade encontrada; quando o mesmo código aparece novamente, a quantidade é somada ao total já registrado.
- Controles para adicionar/remover quantidade manualmente e corrigir erros.
- Suporte inicial a simbologias comuns: EAN-13, EAN-8, UPC-A/UPC-E, Code 128 e QR Code, conforme capacidade da biblioteca/dispositivo.
- Preservar códigos como texto para não perder zeros à esquerda.

### 4. Associação e tratamento de códigos
- Registrar qualquer código como identificador livre, sem procurar correspondência em cadastro.
- Não exigir cadastro nem identificação prévia: cada código é aceito como identificador livre. O operador informa a quantidade encontrada e códigos repetidos acumulam.
- Agrupar códigos iguais e somar as quantidades informadas.
- Não exigir cadastro prévio, descrição ou saldo de estoque.
- Não confundir SKU/código interno com código de barras.

### 5. Sessões e confiabilidade
- Guardar localmente a contagem em andamento e permitir recuperar a sessão ao reimportar o mesmo modelo.
- Salvar cada leitura e alteração de quantidade imediatamente no armazenamento local.
- Funcionar sem internet durante a coleta; sincronização/backup em nuvem pode ser etapa posterior e opcional.
- Recuperar a sessão após fechamento inesperado do aplicativo.
- Registrar horário, código lido, produto associado, quantidade e correções para auditoria.
- Evitar perda de leituras: confirmação visual da gravação local antes de aceitar a próxima operação.

### 6. Conferência antes da exportação
- Resumo de produtos encontrados, coletados, não coletados e códigos desconhecidos.
- Conferência do conjunto de códigos e quantidades informadas.
- Exportar uma linha por código lido, preenchendo as colunas escolhidas de código e quantidade.
- Mostrar claramente a regra de preenchimento antes de gerar o arquivo.

### 7. Exportação no modelo da empresa
- Gerar XLSX com as colunas e cabeçalhos do modelo importado.
- Preservar, sempre que tecnicamente possível, abas, ordem das colunas, cabeçalhos e demais campos que não precisam ser alterados.
- Escrever somente nos campos mapeados pelo usuário.
- Manter códigos como texto e quantidades como números.
- Não prometer preservação de macros, conexões externas ou todos os recursos avançados de qualquer arquivo Excel. Detectar limitações e avisar antes da exportação.
- Disponibilizar prévia e validação do arquivo gerado antes de concluir.

## Regras fundamentais do produto

1. **Independente do MobFlow:** não exige conta, banco de dados ou contratação do MobFlow.
2. **Modelo pertence à empresa:** não forçar um layout único de importação.
3. **Nenhuma leitura pode sumir silenciosamente:** persistir localmente e mostrar falhas.
4. **Agrupamento por código:** códigos repetidos acumulam as quantidades informadas.
5. **Sem cadastro prévio:** qualquer código lido é aceito e armazenado.
6. **Exportação controlada:** nunca sobrescrever o arquivo original; gerar uma cópia de saída.
7. **Transparência:** mostrar o que foi reconhecido, o que será alterado e quais limitações existem.

## Telas previstas

1. **Início:** nova coleta, continuar coleta e histórico.
2. **Importar planilha:** selecionar arquivo e escolher aba.
3. **Mapear colunas:** confirmar código, descrição e quantidade de destino.
4. **Prévia técnica (restrita):** confirmar o mapeamento sem revelar o catálogo ao operador da contagem.
5. **Coletor:** leitura contínua de qualquer código, entrada da quantidade encontrada e agrupamento de códigos repetidos.
6. **Pendências:** códigos desconhecidos, ambíguos ou duplicados na origem.
7. **Resumo:** totais, divergências e validações.
8. **Exportar:** regra de preenchimento, prévia e geração do arquivo.

## Escopo do MVP

A primeira versão deve provar o fluxo completo, sem adicionar complexidade desnecessária:

- Aplicativo Android com interface rápida e adaptada a celular.
- Importação de XLSX e CSV.
- Seleção de aba e mapeamento assistido de colunas.
- Leitura por câmera e entrada por leitor HID.
- Contagem por produto com persistência local/offline.
- Lista de códigos desconhecidos e correção manual.
- Resumo da sessão.
- Exportação XLSX/CSV usando o modelo importado, com cópia de segurança e validação.

## Critérios de aceitação do MVP

- Importar pelo menos modelos de planilha com nomes e ordens de colunas diferentes.
- Permitir corrigir manualmente o mapeamento quando a identificação automática falhar.
- Ler o mesmo código três vezes e registrar exatamente três unidades, sem descartar leituras intencionais.
- Manter a sessão após fechar e reabrir o aplicativo.
- Registrar códigos fora da planilha como pendências, sem perder o código lido.
- Exportar quantidades para as linhas corretas e manter os demais campos relevantes.
- Avisar quando o arquivo tiver recursos que não podem ser preservados com segurança.
- Validar o arquivo exportado reabrindo-o e conferindo os valores escritos.

## Etapas de desenvolvimento

- [ ] **Fase 0 — Fundação:** definir plataforma, arquitetura, biblioteca de leitura e formatos suportados.
- [ ] **Fase 1 — Importador:** abrir XLSX/CSV, listar abas, pré-visualizar e mapear colunas.
- [ ] **Fase 2 — Motor de coleta:** câmera/HID, leitura contínua, contagem e feedback.
- [ ] **Fase 3 — Persistência offline:** sessões, histórico, recuperação e auditoria.
- [ ] **Fase 4 — Conferência:** desconhecidos, ambiguidades, divergências e correções.
- [ ] **Fase 5 — Exportador:** preencher cópia do modelo, validar e disponibilizar arquivo.
- [ ] **Fase 6 — Testes reais:** testar planilhas de empresas diferentes, dispositivos e leitores físicos.

## Fora do primeiro MVP

- Integração direta com ERP de cada empresa.
- Sincronização multiusuário em tempo real.
- Gestão de estoque, vendas ou emissão fiscal.
- Reconhecimento automático garantido de qualquer arquivo complexo sem revisão.
- Painel administrativo SaaS e cobrança por assinatura.

Esses recursos podem ser avaliados depois que o fluxo universal de importação, coleta e exportação estiver confiável.

---

**ColetaFlow — Bipe. Confira. Exporte.**


## Prévia na Vercel

O repositório está conectado ao projeto Vercel `coletaflow`. Commits na branch `main` devem iniciar uma nova implantação automaticamente.
