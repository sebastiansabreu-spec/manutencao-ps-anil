# Sistema Integrado de Gestão Hospitalar - PS ANIL 🏥
**SEMUS - Secretaria Municipal de Saúde de São Luís / Pronto Socorro do Anil**

Sistema completo para substituição de planilhas (*BD_Manutenção*, *BD_Rouparia*, *Lançamentos*, *Consolidado Geral* e *Relatórios*).

---

## 🚀 Como Iniciar

Dê um duplo clique no arquivo:
👉 **`iniciar.bat`**

O servidor será inicializado e abrirá automaticamente no navegador em:
`http://127.0.0.1:8000`

---

## 📑 Módulos e Áreas do Sistema

### 1. ✍️ Painel Central de Lançamentos
- **Módulo / Destino com chaveamento instantâneo**:
  - **Manutenção**: Data, Horário, Turno, Categoria do Serviço, Tipo de Manutenção (Corretiva / Preventiva), Item / Equipamento, Setor / Local, Descrição, Quantidade (Peças/Volume), Nº Chamado/Pedido, Responsável Entrega/Vistoria, Responsável Recebimento/Acompanhamento, Status (Pendente, Solicitado, Ok, Concluído), Data Conclusão/Retorno, Observações e Checkbox de **Retrabalho**.
  - **Rouparia**: Data, Horário, Turno, Tipo de Movimentação (Estoque, Recolhimento Sujo, Entrega aos Setores, Devolução por Não Conformidade), Item Enxoval, Setor/Local, Quantidade (Peças), Peso (KG) / Hampers, Responsável Unidade (Nome e Matrícula), Responsável Empresa Terceirizada (LAVARE) e Inconformidade/Ocorrência.

### 2. 📊 Painel Consolidado de Indicadores
Filtro por competência mensal (`Mês/Ano` ou `Todos`) com as fórmulas exatas validadas:
- **Indicadores de Manutenção**:
  - `Preventivas Realizadas`: Quantidade de preventivas finalizadas com status Concluído ou Ok.
  - `Preventivas Agendadas / Total`: Total de manutenções preventivas do período.
  - `% Aderência Preventiva`: `(Preventivas Realizadas / Preventivas Agendadas) * 100`.
  - `Corretivas Solicitadas`: Total de corretivas abertas no período.
  - `Corretivas Concluídas`: Corretivas finalizadas com status Concluído ou Ok.
  - `% Resolução de Corretivas`: `(Corretivas Concluídas / Corretivas Solicitadas) * 100`.
  - `Total Peças / Insumos Aplicados`: Soma da coluna Quantidade (Peças/Volume).
  - `Taxa de Retrabalho`: `(Contagem de Retrabalho / Total de Concluídas) * 100`.
- **Indicadores de Rouparia**:
  - `Total de Peças Entregues aos Setores`: Soma de peças distribuídas.
  - `Total de Enxoval Sujo Recolhido (KG)`: Soma do peso recolhido.
  - `Total de Enxoval Limpo Recebido (KG)`: Soma do peso recebido/armazenado.
  - `Devoluções por Não Conformidade (KG / Peças)`: Peças/KG reprovadas.
  - `Nível de Atendimento de Rouparia (%)`: `(Entregues / (Entregues + Devoluções)) * 100`.
  - `Pontualidade na Entrega (07h às 19h)`: % de entregas dentro da faixa de horário oficial.
  - `Custo Total Lavanderia (R$)`: `Total KG Sujo * Custo R$/KG`.
  - `Custo por Quilo (R$/KG)`: Parâmetro cadastrado nas configurações.

### 3. 📈 Relatórios & Gráficos
- **Gráfico de Rosca / Pizza**: Distribuição por Status das Manutenções.
- **Gráfico de Barras**: Manutenções agrupadas por Categoria.
- **Ranking por Setor / Local**: Identificação dos setores com maior volume de chamados.
- **Matriz Cruzada (Setor × Item / Equipamento)**: Grade bidimensional exata com a contagem cruzada de ocorrências e insumos.

### 4. 🗄️ Base de Dados (Tabelas)
- Visualização e busca rápida nas bases `BD_Manutencao` e `BD_Rouparia`.
- Histórico com IDs oficiais (ex: `PRED-0001`, `CLIM-0003`, `R-9`, etc.).

### 5. ⚙️ Configurações & Cadastros Base
- Inclusão e exclusão de novos **Itens & Insumos de Manutenção**.
- Inclusão e exclusão de novos **Itens de Enxoval (Rouparia)**.
- Inclusão e exclusão de **Setores / Locais do Hospital**.
- Inclusão e exclusão de **Categorias de Manutenção**.
- Configuração do parâmetro de **Custo por Quilo da Lavanderia terceirizada (R$/KG)**.

### 6. 📥 Exportação de Relatórios
- **Consolidado Geral (.xlsx)**: Planilha estruturada idêntica ao painel impresso/oficial.
- **BD Manutenção (.xlsx)**: Base de dados completa com todas as 17 colunas.
