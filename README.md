Flask
waitress
openpyxl
pypdf
Werkzeug
psycopg2-binary

#  Sistema Inteligente de Gestão de Resíduos Hospitalares

Sistema web desenvolvido em **Python (Flask)** para o lançamento, classificação, monitoramento de pesagem e geração de relatórios de resíduos sólidos hospitalares. O projeto conta com painéis dinâmicos, gráficos interativos e suporte a importação de relatórios em PDF e exportação para Excel.

---

##  Aceder à Aplicação Online (Deploy)
A aplicação encontra-se em funcionamento na nuvem através do Render:
 **[Aceder ao Sistema de Gestão de Resíduos](https://controle-de-residuos-hospital.onrender.com/)**

---

##  Tecnologias Utilizadas
O projeto foi construído utilizando as seguintes tecnologias e bibliotecas:
* **Backend:** Python, Flask, Waitress (Servidor WSGI de produção)
* **Base de Dados:** PostgreSQL (Produção na nuvem) e SQLite (Desenvolvimento local)
* **Frontend & UI:** HTML5, Tailwind CSS, Lucide Icons, Chart.js (Gráficos interativos)
* **Manipulação de Dados & Ficheiros:** 
  * `openpyxl` (Geração e estilização avançada de planilhas Excel)
  * `pypdf` (Extração e leitura automatizada de dados de relatórios em PDF)
  * `Werkzeug` (Segurança e encriptação de palavras-passe)
  * `psycopg2-binary` (Conector para PostgreSQL)

---

##  Funcionalidades Principais
1. **Autenticação Segura:** Sistema de login com gestão de perfis de utilizador (Administrador e Operador).
2. **Lançamento Manual e Automático:** Registo de pesagens por setor/localidade (Infectante, Compostagem, Comum, Reciclagem e Químico) ou importação direta via PDF.
3. **Dashboards Dinâmicos:** Gráficos mensais consolidados e distribuição percentual por setor com exportação rápida para PNG.
4. **Relatórios Personalizados:** Geração de planilhas em Excel (`.xlsx`) com filtros avançados por ano e setor.
5. **Gestão Administrativa:** Módulo exclusivo para administradores gerirem utilizadores e histórico de registos.

---

##  Como Executar o Projeto Localmente

1. Clone o repositório:
   ```bash
   git clone https://github.com/Philipe229/controle-de-residuos-hospital.git
