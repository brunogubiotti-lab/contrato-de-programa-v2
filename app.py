"""
app.py — SISPLAN LIGHT (versão de demonstração)
==================================================
Versão enxuta do SISPLAN, para apresentação. Diferenças em relação
ao sistema real:
  - SEM login (Flask-Login removido) — acesso livre de propósito
  - SEM módulos de exportação (AGEMS, SINISA) — só visualização
  - SEM banco de dados de usuários (SQLite/SQLAlchemy removidos)

Como rodar localmente:
    pip install -r requirements.txt
    python app.py

Em produção (Render.com):
    gunicorn app:app
"""

from flask import Flask


def create_app():
    app = Flask(__name__)

    from blueprints.home import home_bp
    from blueprints.contratos import contratos_bp

    app.register_blueprint(home_bp)
    app.register_blueprint(contratos_bp)

    return app


app = create_app()

if __name__ == "__main__":
    app.run(debug=True, port=5000)
