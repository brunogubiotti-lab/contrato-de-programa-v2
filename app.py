"""
app.py — SISPLAN LIGHT (versão de demonstração)
==================================================
Versão enxuta do SISPLAN, para apresentação. Diferenças em relação
ao sistema real:
  - SEM login (Flask-Login removido) — acesso livre de propósito
  - SEM banco de dados de usuários (SQLite/SQLAlchemy removidos)

Como rodar localmente:
    pip install -r requirements.txt
    python app.py

Em produção (Render.com):
    gunicorn app:app
"""

from flask import Flask


def _brnum(value):
    """Formata números no padrão brasileiro: milhar com ponto, decimal
    com vírgula. Strings e None passam direto (usado no template do RAD,
    onde nem todo campo é numérico — ex.: nomes de ETE, Captação)."""
    if value is None or isinstance(value, str):
        return value
    try:
        numero = float(value)
    except (TypeError, ValueError):
        return value
    if numero == int(numero):
        texto = f"{int(numero):,}".replace(",", ".")
    else:
        texto = f"{numero:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return texto


def create_app():
    app = Flask(__name__)
    app.jinja_env.filters["brnum"] = _brnum

    from blueprints.home import home_bp
    from blueprints.contratos import contratos_bp
    from blueprints.relatorios_agems import relatorios_agems_bp

    app.register_blueprint(home_bp)
    app.register_blueprint(contratos_bp)
    app.register_blueprint(relatorios_agems_bp)

    return app


app = create_app()

if __name__ == "__main__":
    app.run(debug=True, port=5000)
