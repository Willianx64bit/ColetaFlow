# ColetaFlow

**Coletor inteligente universal de produtos.**


## Estado do desenvolvimento

**MVP inicial com fluxo de bipagem cega implementado — aguardando validação com planilhas reais.**

Já existe uma primeira interface React + TypeScript com:
- Importação de arquivos XLSX, XLS e CSV.
- Sugestão automática e mapeamento manual das colunas de código, descrição e quantidade.
- Entrada por leitor físico em modo teclado (HID).
- Leitura pela câmera usando ZXing.
- Bipagem cega: durante a coleta, a interface não exibe nomes, descrições, catálogo completo nem saldo esperado; mostra somente códigos já bipados e quantidades.
- Contagem por código, lista de códigos não encontrados e alerta para códigos duplicados na planilha.
- Ajuste manual das quantidades e exportação para XLSX.
- Códigos desconhecidos acrescentados ao final da planilha exportada; códigos ambíguos ficam pendentes.
- Contagem guardada localmente no navegador e recuperável ao reimportar o mesmo arquivo.
- Workflow do GitHub Actions para verificar instalação e build.

**Limitação atual:** ainda é uma primeira base de desenvolvimento, não uma versão validada para operação comercial. O modo cego foi aplicado à tela de coleta e à lista de conferência; a etapa de importação/mapeamento ainda precisa ser protegida caso o mesmo dispositivo seja entregue diretamente ao operador. A leitura da câmera exige HTTPS e permissão do dispositivo. A sessão guarda as contagens localmente, mas o arquivo de origem precisa ser reimportado após recarregar a página. Preservação avançada de macros, fórmulas, estilos especiais e layouts complexos ainda precisa de testes.

Próximos passos: validar o build, testar com planilhas reais variadas, reforçar a persistência offline e implementar instalação PWA/Android.

O ColetaFlow é um aplicativo independente para coletar códigos de barras e quantidades de produtos e devolver os resultados no modelo de planilha fornecido por cada empresa.

> **Princípio do produto:** a empresa mantém seu próprio modelo; o ColetaFlow adapta a coleta e a exportação a esse modelo, sem exigir uma planilha padrão única.

## Objetivo

Permitir que uma pessoa ou equipe importe uma planilha de produtos, faça a contagem por leitura de código de barras e exporte o resultado preenchido no formato esperado pela empresa.

Fluxo principal:

**Importar planilha → interpretar colunas → revisar mapeamento → coletar/bipar → conferir → exportar no modelo original.**

## Funcionalidades planejadas

### 1. Importação de arquivos
- Primeira versão: arquivos Excel `.xlsx` e CSV.
- Identificar abas, cabeçalhos e linhas de produtos.
- Reconhecer possíveis colunas de código de barras, SKU/código interno, descrição e quantidade.
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
- Cada leitura válida soma uma unidade por padrão; leitura repetida é permitida e representa outra unidade.
- Controles para adicionar/remover quantidade manualmente e corrigir erros.
- Suporte inicial a simbologias comuns: EAN-13, EAN-8, UPC-A/UPC-E, Code 128 e QR Code, conforme capacidade da biblioteca/dispositivo.
- Preservar códigos como texto para não perder zeros à esquerda.

### 4. Associação e tratamento de códigos
- Associar a leitura à linha correta usando o identificador configurado.
- **Modo cego como padrão:** não exibir descrição, nome, preço, catálogo completo ou saldo esperado durante a coleta. A interface de conferência mostra somente código e quantidade.
- Separar códigos não encontrados em uma lista de pendências, sem descartá-los.
- Tratar códigos duplicados na planilha com uma regra explícita: sinalizar e permitir escolher a linha ou consolidar somente após confirmação.
- Não confundir SKU/código interno com código de barras.

### 5. Sessões e confiabilidade
- Criar, pausar, retomar, finalizar e consultar sessões de coleta.
- Salvar cada leitura e alteração de quantidade imediatamente no armazenamento local.
- Funcionar sem internet durante a coleta; sincronização/backup em nuvem pode ser etapa posterior e opcional.
- Recuperar a sessão após fechamento inesperado do aplicativo.
- Registrar horário, código lido, produto associado, quantidade e correções para auditoria.
- Evitar perda de leituras: confirmação visual da gravação local antes de aceitar a próxima operação.

### 6. Conferência antes da exportação
- Resumo de produtos encontrados, coletados, não coletados e códigos desconhecidos.
- Comparação entre quantidade original e quantidade coletada, quando houver quantidade-base.
- Permitir exportar quantidade coletada, substituir quantidade existente ou somar à quantidade existente, conforme regra escolhida pelo usuário.
- Mostrar claramente a regra de preenchimento antes de gerar o arquivo.

### 7. Exportação no modelo da empresa
- Gerar XLSX ou CSV com as colunas e cabeçalhos esperados pelo modelo importado.
- Preservar, sempre que tecnicamente possível, abas, ordem das colunas, identificadores, descrições e demais campos que não precisam ser alterados.
- Escrever somente nos campos mapeados pelo usuário.
- Manter códigos como texto e quantidades como números.
- Não prometer preservação de macros, conexões externas ou todos os recursos avançados de qualquer arquivo Excel. Detectar limitações e avisar antes da exportação.
- Disponibilizar prévia e validação do arquivo gerado antes de concluir.

## Regras fundamentais do produto

1. **Independente do MobFlow:** não exige conta, banco de dados ou contratação do MobFlow.
2. **Modelo pertence à empresa:** não forçar um layout único de importação.
3. **Nenhuma leitura pode sumir silenciosamente:** persistir localmente e mostrar falhas.
4. **Sem deduplicação indevida:** duas bipagens intencionais do mesmo produto devem contar duas unidades.
5. **Sem associação insegura:** códigos desconhecidos ou ambíguos ficam pendentes.
6. **Exportação controlada:** nunca sobrescrever o arquivo original; gerar uma cópia de saída.
7. **Transparência:** mostrar o que foi reconhecido, o que será alterado e quais limitações existem.

## Telas previstas

1. **Início:** nova coleta, continuar coleta e histórico.
2. **Importar planilha:** selecionar arquivo e escolher aba.
3. **Mapear colunas:** confirmar código, descrição e quantidade de destino.
4. **Prévia técnica (restrita):** confirmar o mapeamento sem revelar o catálogo ao operador da contagem.
5. **Coletor cego:** leitura contínua, código e quantidade, sem descrição ou saldo esperado.
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
