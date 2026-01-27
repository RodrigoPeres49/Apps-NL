import pandas as pd
import os


caminho_pendencias = r"C:\Users\rodrigo.souza\nlfrutas.com.br\NL_File-Server - Documentos\Comercial_Terceirizadas\CONSOLIDA_VENDAS\Vendas"
controle_pendencias = "Vendas Terceirizadas Lançamentos.xlsx"

pendencias = pd.read_excel(os.path.join(caminho_pendencias, controle_pendencias), sheet_name="Pendencias")
pendencias.columns = pendencias.columns.str.strip()

pendencias["Data da Venda"] = pd.to_datetime(pendencias["Data da Venda"], errors="coerce")
pendencias = pendencias.drop("Pendencias", axis=1)
pendencias = pendencias.rename(columns={"Data da Venda": "Dias Pendentes"})

filtros = ~(

    (
        (pendencias["Rede"] == "BADIAO") & (
            (
                (pendencias["Dias Pendentes"] < pd.to_datetime("2026-01-22")) 
            )
        )
    ) |

    (
        (pendencias["Rede"] == "DONADIO") & (
            (
                (pendencias["Dias Pendentes"] == pd.to_datetime("2025-12-14")) 
            )
        )
    ) |

    (pendencias["Dias Pendentes"] == pd.to_datetime("2025-12-25")) |

    (
        (pendencias["Rede"] == "CONSUL") & (
            (
                (pendencias["Dias Pendentes"] == pd.to_datetime("2025-11-20")) 
            )
        )
    ) |

    (
        (pendencias["Rede"] == "BRAGA") & (
            (
                (pendencias["Dias Pendentes"] == pd.to_datetime("2025-10-12")) 
            )
        )
    ) |

    (
        (pendencias["Rede"] == "CONSUL") & (
            (
                (pendencias["Dias Pendentes"] == pd.to_datetime("2025-10-12")) 
            )
        )
    ) |

    (
        (pendencias["Rede"].isin(["PAG POUCO", "ESCOLA"])) &
        (pendencias["Dias Pendentes"].dt.weekday == 6)
    )
)

pendencias = pendencias[filtros]
pendencias["Dias Pendentes"] = pendencias["Dias Pendentes"].dt.strftime("%d/%m/%Y")
pendencias = pendencias.to_dict(orient="records")



