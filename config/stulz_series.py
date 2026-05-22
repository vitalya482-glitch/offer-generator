from __future__ import annotations

# Mapping of STULZ model-code prefixes to product-line metadata used by
# brands/stulz.py when composing {{intro_text}}.
#
# Add new prefixes here instead of changing the offer-generation logic.
# Prefix matching is case-insensitive and uses the first letter/digit block
# before separators such as '-', space, '/', etc.

STULZ_SERIES = {
    # CyberAir
    "ASU": {
        "line": "Stulz CyberAir",
        "equipment_type_single": "прецизионного кондиционера",
        "equipment_type_plural": "прецизионных кондиционеров",
        "install_type": "напольного исполнения",
        "airflow": "upflow",
    },
    "ASD": {
        "line": "Stulz CyberAir",
        "equipment_type_single": "прецизионного кондиционера",
        "equipment_type_plural": "прецизионных кондиционеров",
        "install_type": "напольного исполнения",
        "airflow": "downflow",
    },

    # Telecom
    "SXL": {
        "line": "Stulz Telecom",
        "equipment_type_single": "телекоммуникационного кондиционера",
        "equipment_type_plural": "телекоммуникационных кондиционеров",
        "install_type": "напольного исполнения",
        "airflow": "upflow",
    },
}

DEFAULT_STULZ_SERIES = {
    "line": "Stulz",
    "equipment_type_single": "прецизионного кондиционера",
    "equipment_type_plural": "прецизионных кондиционеров",
    "install_type": "напольного исполнения",
    "airflow": "",
}

AIRFLOW_TEXT = {
    "upflow": "с верхней подачей охлажденного воздуха",
    "downflow": "с нижней подачей охлажденного воздуха под фальшпол",
    "front": "с фронтальной подачей охлажденного воздуха",
    "front_duct": "с фронтальной подачей охлажденного воздуха через короб",
    "horizontal": "с горизонтальной подачей охлажденного воздуха",
}
