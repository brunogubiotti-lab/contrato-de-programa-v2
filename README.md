# SISPLAN — Versão Light (Demonstração)

Versão reduzida do SISPLAN para apresentação, sem login e sem exportação
de relatórios. Contém apenas o módulo **Contratos de Programa**:
comparação Real (dados SINISA) x Meta contratual, por município e por
indicador, para os 68 municípios atendidos pela SANESUL.

## O que NÃO está aqui (por decisão, não por limitação técnica)
- Login / usuários (acesso livre de propósito, é só demonstração)
- Módulo AGEMS (exportação de relatórios em Excel/PDF)
- Módulo SINISA (glossários, validação)
- Painel (Rel. Gerencial, RAD, Investimentos, NR-08, NR-09) — pode ser
  adicionado depois, mas precisa do código-fonte de `painel.py` para
  ser portado com segurança (não incluído nesta rodada).
- Tela de Projeções (ainda não implementada no sistema real também)

## Como rodar localmente
```bash
pip install -r requirements.txt
python app.py
```
Abre em `http://localhost:5000`

## Deploy no Render
1. Suba esta pasta inteira para um repositório novo no GitHub (separado
   do repositório do sistema real).
2. No Render: **New → Web Service** → conecte o repositório.
3. Configurações:
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app` (já está no Procfile, o Render
     detecta sozinho na maioria dos casos)
   - **Instance Type**: Free
4. Não precisa configurar nenhuma variável de ambiente — essa versão
   não usa SECRET_KEY, banco de dados ou e-mail.
5. Deploy. Em alguns minutos a URL pública estará no ar.

## Sobre os dados
`data/base_dados_sisplan.xlsx` é UMA planilha só, com duas abas:
- `dados_reais`: um município por linha (valores reais do SINISA já calculados)
- `metas_contrato`: uma linha por (município, indicador), com a meta contratual

Se quiser atualizar os números antes de apresentar, gere essa planilha
de novo a partir do sistema real (mesma lógica dos loaders originais
em `scripts/contratos_loader.py`, `ipl_loader.py` e `metas_loader.py`
do projeto principal) e substitua o arquivo aqui — mesmas duas abas,
mesmas colunas. Não precisa mexer em nenhum código.
