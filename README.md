# Dashboard de Chamados SIMER

📊 Dashboard Interativo

<p align="center">
  <img src="https://img.shields.io/badge/STATUS-EM%20DESENVOLVIMENTO-blue"/>
  <img src="https://img.shields.io/badge/VERSÃO-1.0.0-green"/>
  <img src="https://img.shields.io/badge/LICENSE-MIT-yellow"/>
</p>

## 🧠 Sobre o Projeto

Dashboard interativo desenvolvido para visualização, análise e acompanhamento de dados de chamados de forma clara, rápida e eficiente. O projeto centraliza informações importantes e auxilia na tomada de decisões através de gráficos, KPIs e relatórios.

## 🚀 Funcionalidades

- ✅ Visualização de dados em tempo real
- ✅ Gráficos interativos com Chart.js
- ✅ Indicadores de performance (KPIs)
- ✅ Filtros personalizados
- ✅ Interface responsiva
- ✅ Suporte a temas claro/escuro
- ✅ Exportação de dados filtrados
- ✅ Carregamento de planilhas Excel
- ✅ Comparação de arquivos Excel

## 🛠️ Tecnologias Utilizadas

- **Frontend:** HTML5, CSS3, JavaScript (Vanilla)
- **Bibliotecas:** Chart.js, XLSX (SheetJS)
- **Estrutura:** Modular com arquivos utilitários compartilhados

## 📁 Estrutura do Projeto

```
dashboard_chamados_simer/
├── index.html                 # Página principal com navegação
├── dashboard.html             # Dashboard principal
├── Paginas/
│   └── dashboard_ocorrencias_erros_operacionais.html
├── assets/
│   ├── css/
│   │   ├── portal.css
│   │   ├── ocorrencias.css
│   │   └── erros_operacionais.css
│   ├── js/
│   │   ├── utils.js          # Funções utilitárias compartilhadas
│   │   ├── portal.js         # Navegação e abas
│   │   ├── ocorrencias.js    # Lógica do dashboard de ocorrências
│   │   └── erros_operacionais.js # Lógica de erros operacionais
│   └── img/
│       └── icon.png
└── README.md
```

## 🔧 Como Executar

### Opção recomendada (com backend Node.js)

1. Clone o repositório
2. Instale as dependências:
   - `npm install`
3. Inicie o servidor:
   - `npm run dev`
4. Abra no navegador:
   - `http://localhost:3000`

### Opção antiga (apenas arquivos estáticos)

Ainda funciona abrir o `index.html` diretamente, mas o fluxo recomendado é via Node para permitir o parse no backend e evitar diferenças de ambiente.

## 📈 Refatoração Realizada

- **Modularização:** Criado `utils.js` com funções compartilhadas (`$`, `normKey`, `findValue`, etc.)
- **Eliminação de código duplicado:** Removidas funções repetidas entre arquivos JS
- **Padronização:** Consistência nos caminhos de assets
- **Manutenibilidade:** Código mais organizado e reutilizável

## 👨‍💻 Autor

Desenvolvido por Vicente Freitas

- 🔗 GitHub: https://github.com/vicenteVJ
- 🔗 LinkedIn: https://www.linkedin.com/in/vicente-joel-096829148/
