from automation_functions import *

# ============================================================

# CONSULTAR NOTA

# ============================================================

def consultar_nota(
numero_nota
):


    print()
    print("================================================")
    print("INICIANDO CONSULTA DA NOTA")
    print("================================================")
    
    print(
        "NÚMERO DA NOTA:",
        numero_nota
    )
    
    
    try:
    
        # ====================================================
        # CONECTAR AO CHROME
        # ====================================================
    
        driver = conectar_chrome()
    
    
        xml = mudar_aba(
            driver,
            "Portal de importação de XML"
        )
        
        if not xml:
        
            return {
                "sucesso": False,
                "mensagem": (
                    "Portal de importação de XML "
                    "não encontrado."
                )
            }
        
        driver.switch_to.default_content()
    
    
        # ====================================================
        # 1 - CLICAR NO PRIMEIRO APLICAR
        # ====================================================
    
        seletor = (
            "button"
            "[ng-click='applyFilter();']"
        )
    
    
        print()
        print("================================================")
        print("1 - CLICANDO NO PRIMEIRO APLICAR")
        print("================================================")
    
    
        sucesso = clicar_campo(
            driver,
            seletor
        )
    
    
        if not sucesso:
    
            return {
                "sucesso": False,
                "mensagem": (
                    "Não foi possível clicar "
                    "no primeiro botão Aplicar."
                )
            }
    
    
        # ====================================================
        # AGUARDAR FORMULÁRIO DE CONSULTA
        # ====================================================
    
        time.sleep(
            1
        )
    
    
        # ====================================================
        # 2 - PREENCHER NÚMERO DA NOTA
        # ====================================================
    
        seletor = (
            "input"
            "[type='text']"
            "[ng-model='value']"
            "[ng-keydown='onInputKeyDown($event)']"
            "[ng-change='changeHandler()']"
        )
    
    
        print()
        print("================================================")
        print("2 - PREENCHENDO NÚMERO DA NOTA")
        print("================================================")
    
    
        sucesso = preencher_campo(
            driver,
            seletor,
            numero_nota
        )
    
    
        if not sucesso:
    
            return {
                "sucesso": False,
                "mensagem": (
                    "Não foi possível preencher "
                    "o número da nota."
                )
            }
    
    
        # ====================================================
        # 3 - CLICAR NO SEGUNDO APLICAR
        # ====================================================
    
        seletor = (
            "button"
            "[ng-click='onSuccess()']"
        )
    
    
        print()
        print("================================================")
        print("3 - CLICANDO NO SEGUNDO APLICAR")
        print("================================================")
    
    
        sucesso = clicar_campo(
            driver,
            seletor
        )
    
    
        if not sucesso:
    
            return {
                "sucesso": False,
                "mensagem": (
                    "Não foi possível clicar "
                    "no segundo botão Aplicar."
                )
            }
            

        # ====================================================
        # 4 - CLICAR NO CÓDIGO DO TIPO DE OPERAÇÃO
        # ====================================================
        
        print()
        print("================================================")
        print("4 - CLICANDO NO CÓDIGO DO TIPO DE OPERAÇÃO")
        print("================================================")
        
        
        seletor = (
            "div"
            "[role='gridcell']"
            "[col-id='CODTIPOPER']"
        )
        
        
        sucesso = clicar_campo(
            driver,
            seletor
        )
        
        
        if not sucesso:
        
            return {
                "sucesso": False,
                "mensagem": (
                    "Não foi possível clicar "
                    "no código do tipo de operação."
                )
            }
            

        # 5 - preencher 2200
        print()
        print("================================================")
        print("5 - PREENCHENDO CÓDIGO DO TIPO DE OPERAÇÃO")
        print("================================================")

        seletor = (
            "div"
            "[role='gridcell']"
            "[col-id='CODTIPOPER']"
            "[class*='ag-cell-inline-editing']"
            " sk-pesquisa-input"
            "[sk-target-field='CODTIPOPER']"
            " sk-text-input"
            "[sk-value='codeValue']"
            " input"
            "[type='text']"
        )

        sucesso = digitar_no_foco(
            driver,
            "2200"
        )

        if not sucesso:
            return {
                "sucesso": False,
                "mensagem": (
                    "Não foi possível preencher "
                    "o código 2200."
                )
            }



        # ====================================================
        # FINAL
        # ====================================================
    
        print()
        print("================================================")
        print("CONSULTA REALIZADA COM SUCESSO")
        print("================================================")
    
    
        return {
            "sucesso": True,
            "mensagem": (
                f"Consulta da nota {numero_nota} "
                "realizada com sucesso."
            )
        }
    
    
    except Exception as erro:
    
        print()
        print("================================================")
        print("ERRO NA AUTOMAÇÃO")
        print("================================================")
    
        print(
            str(erro)
        )
    
    
        return {
            "sucesso": False,
            "mensagem": str(erro)
        }


