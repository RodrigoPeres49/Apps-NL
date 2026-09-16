from automations.automation_functions import *
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from datetime import datetime

# ============================================================

# CONSULTAR NOTA

# ============================================================

def consultar_nota_devolucao(numero_nota, data, tipo, motivo):

    print("INICIANDO LANÇAMENTO NÚMERO DA NOTA:",numero_nota)

    # CORRIGINDO FORMATO DA DATA

    data = datetime.strptime(data, "%Y-%m-%d")
    data_lançamento = data.strftime("%d/%m/%Y")


    try:
    
        # ====================================================
        # CONECTAR AO CHROME
        # ====================================================
    
        driver = conectar_chrome()
        xml = mudar_aba(driver,"Portal de importação de XML")
        
        if not xml:
        
            return {
                "sucesso": False,
                "mensagem": (
                    "Portal de importação de XML "
                    "não encontrado."
                )
            }
        
        driver.switch_to.default_content()

        seletor_tela = (
        "button"
        "[ng-click='ctrl.changeFace()']"
        )

        tela = verificar_tela(driver)
        if tela == "Inicial XML":
            pass
        elif tela == "Grade XML":
            clicar_campo(driver, seletor_tela)
            time.sleep(1) 

        else:
            return{
                "sucesso": False,
                "mensagem":"Não foi possível encontrar o portal de importação de XML"
            }
        # ====================================================
        # 1 - CLICAR NO PRIMEIRO APLICAR
        # ====================================================
    
        seletor = (
            "button"
            "[ng-click='applyFilter();']"
        )
      
        sucesso = clicar_campo(driver,seletor)
        time.sleep(1)
    
    
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
    
        time.sleep(1)
    
    
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

        time.sleep(8)  

        # ====================================================
        # 4 - CLICAR NO CÓDIGO DO TIPO DE OPERAÇÃO
        # ====================================================
        
        
        seletor = (
            "div"
            "[role='gridcell']"
            "[col-id='CODTIPOPER']"
        )
        
        
        sucesso = clicar_campo(driver,seletor)
        
        
        if not sucesso:
        
            return {
                "sucesso": False,
                "mensagem": (
                    "Não foi possível clicar "
                    "no código do tipo de operação."
                )
            }

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


        time.sleep(2)    

        # 5 - PREENCHER TIPO DE NOTA 

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

        sucesso = digitar_no_foco(driver,tipo)
        # SALVANDO
        ActionChains(driver).send_keys(Keys.F7).perform()

        if not sucesso:
            return {
                "sucesso": False,
                "mensagem": (
                    "Não foi possível preencher "
                    "o código.", tipo
                )
            }

        time.sleep(2)

        seletor = (
            "button"
            "[ng-click='ctrl.processarArquivo()']"
        )

        sucesso = clicar_campo(driver, seletor)

        if not sucesso:

            return {
                "sucesso": False,
                "mensagem": (
                    "Não foi possível clicar "
                    "Em processar arquivos."
                )
            }


        seletor = "h5.modal-title"
        
        while True:
        
            if verificar_elemento(driver,seletor,"Aviso"):
                break

            if verificar_elemento(driver,seletor, "Erro"):

                print("GATEWAY TIME-OUT FAVOR VERIFICAR")
                time.sleep(180)

                seletor = (
                    "button"
                    "[ng-click='applyFilter();']"
                )

                clicar_campo(driver,seletor)

                time.sleep(1)

                seletor = (
                    "button"
                    "[ng-click='onSuccess()']"
                )

                clicar_campo(driver,seletor)
                break

            print("\rAGUARDANDO PROCESSAR ARQUIVOS...",end="",flush=True)
            time.sleep(1)


        time.sleep(2)
    
        seletor = ("button""[ng-click='onSuccess()']")

        sucesso = clicar_campo(driver, seletor)

        time.sleep(2)

            

        # ============================================================
        # VERIFICAR DIVERGENCIAS
        # ============================================================
    
        seletor = (
            "button"
            "[ng-click='ctrl.changeFace()']"
            )
        
        clicar_campo(driver, seletor)
        time.sleep(1.5)


            
        resultado_divergencias = verificar_divergencias_devolucao(driver)

        tela = ""
        tela = verificar_tela(driver)
        time.sleep(1)

        if tela == "Grade XML":
            pass
        elif tela == "Inicial XML":
            clicar_campo(driver, seletor)
        else:
            return {
                "sucesso": False,
                "mensagem": "Grade XML não encontrada."
            }
        

            
        if resultado_divergencias["existe"]:
            print("Divergência encontrada!")
            time.sleep(1)
        
            # ============================================================
            # TRATAR TODAS AS DIVERGÊNCIAS
            # ============================================================
        
            precisa_recalcular = False
        
            for divergencia in resultado_divergencias["divergencias"]:
        
                tipo_divergencia = divergencia["tipo"]
        
                # ========================================================
                # FINANCEIRO
                # ========================================================
        
                if tipo_divergencia == "Financeiro":
        
                    seletor_financeiro = (
                        "//div[@flex and @layout='row' and @ng-click='select()']"
                        "//sk-i18n[normalize-space()='Financeiro']"
                    )
        
                    sucesso = clicar_campo(
                        driver,
                        seletor_financeiro
                    )
        
                    if not sucesso:
                        return {
                            "sucesso": False,
                            "mensagem": (
                                "Não foi possível clicar na aba Financeiro."
                            )
                        }
        
                    time.sleep(3)
        
                    # ----------------------------------------------------
                    # CLICAR EM NÃO INFORMADO
                    # ----------------------------------------------------
        
                    seletor_financeiro_nao_informado = (

                        "div.ui-select-match[title='Usar Financeiro do Sistema'], "
                        "div.ui-select-match[title='Não Informado']"
                    )
                    

                    time.sleep(3)
        
                    sucesso = clicar_campo(
                        driver,
                        seletor_financeiro_nao_informado
                    )
        
                    if not sucesso:

                        return{
                            "sucesso": False,
                            "mensagem": "Não foi possível clicar em Usar Financeiro do Sistema"
                        }

        
        
                    # ----------------------------------------------------
                    # USAR FINANCEIRO DO SISTEMA
                    # ----------------------------------------------------
                    seletor_usar_financeiro = (
                        "//div[contains(@class,'option') "
                        "and contains(@class,'ui-select-choices-row-inner')]"
                        "//span[normalize-space()='Usar Financeiro do Sistema']"
                    )
        
                    sucesso = clicar_campo(driver,seletor_usar_financeiro)
                    time.sleep(1)
        
                    if not sucesso:
                        return {
                            "sucesso": False,
                            "mensagem": (
                                "Não foi possível selecionar "
                                "Usar Financeiro do Sistema."
                            )
                        }
        
                    print("Financeiro configurado com sucesso.")
        
                    precisa_recalcular = True
        
                    time.sleep(2)
        
        
                # ========================================================
                # IMPOSTOS
                # ========================================================
        
                elif tipo_divergencia == "Impostos":
        
                    seletor_impostos = (
                        "//div[@flex and @layout='row' and @ng-click='select()']"
                        "//sk-i18n[normalize-space()='Impostos']"
                    )
        
                    sucesso = clicar_campo(driver,seletor_impostos)
        
                    if not sucesso:
                        return {
                            "sucesso": False,
                            "mensagem": (
                                "Não foi possível clicar na aba Impostos."
                            )
                        }
        
                    time.sleep(2)
        
                    # ----------------------------------------------------
                    # CLICAR EM NÃO INFORMADO
                    # ----------------------------------------------------
        
                    seletor_impostos_nao_informado = (
                        "div.ui-select-match[title='Usar Financeiro do Sistema'], "
                        "div.ui-select-match[title='Não Informado']"

                    )
        
                    sucesso = clicar_campo(driver,seletor_impostos_nao_informado)
        
                    if not sucesso:
                            return{
                                "sucesso": False,
                                "mensagem": "Não foi possível clicar em Usar Impostos do Arquivo"
                            }
        
                    time.sleep(2)
        
                    # ----------------------------------------------------
                    # USAR IMPOSTOS DO ARQUIVO
                    # ----------------------------------------------------
        
                    seletor_usar_impostos = (
                        "//div[contains(@class,'option') "
                        "and contains(@class,'ui-select-choices-row-inner')]"
                        "//span[normalize-space()='Usar Impostos do Arquivo']"
                    )
        
                    sucesso = clicar_campo(driver,seletor_usar_impostos)
        
                    if not sucesso:
                        return {
                            "sucesso": False,
                            "mensagem": (
                                "Não foi possível selecionar "
                                "Usar Impostos do Arquivo."
                            )
                        }
        
                    print("Impostos configurados com sucesso.")
        
                    precisa_recalcular = True
        
                    time.sleep(2)
        
        
                # ========================================================
                # PRODUTOS NÃO ENCONTRADOS
                # ========================================================
        
                elif tipo_divergencia == "Produtos não encontrados":
        
                    return{
                        "sucesso": False,
                        "mensagem": "PRODUTOS NÃO ENCONTRADOS"
                                    "FAVOR REALIZAR MANUALMENTE "
                                    "A INSERÇÃO DOS PRODUTOS"
                    }
        
        
                # ========================================================
                # LOTES
                # ========================================================
        
                elif tipo_divergencia == "Lotes de produtos não informados":

                    return{
                        "sucesso": False,
                        "mensagem": "LOTES DE PRODUTOS NÃO INFORMADOS "
                                    "FAVOR REALIZAR MANUALMENTE "
                                    "A INSERÇÃO DOS PRODUTOS"
                    }
        
        
                # ========================================================
                # OUTRA DIVERGÊNCIA
                # ========================================================
        
                else:
        
                    print()
                    print("Outra divergência encontrada:")
                    print(tipo_divergencia)
        
        
            # ============================================================
            # DEPOIS DE TRATAR TODAS AS DIVERGÊNCIAS
            # FAZER O RECÁLCULO APENAS UMA VEZ
            # ============================================================
        
            if precisa_recalcular:
        
                print()
                print("VALIDANDO IMPORTAÇÃO")
        
                seletor = (
                    "button"
                    "[ng-click='ctrl.recalcular()']"
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
                            "em Validar Importação."
                        )
                    }
        
                # ========================================================
                # AGUARDAR AVISO / INFORMAÇÃO / ERRO
                # ========================================================
        
                seletor = "h5.modal-title"
        
                while True:
        
                    if (
                        verificar_elemento(driver, seletor, "Aviso")
                        or
                        verificar_elemento(driver, seletor, "Informação")
                    ):
        
                        print("PROCESSAMENTO CONCLUÍDO")
        
                        break
        
        
                    if verificar_elemento(
                        driver,
                        seletor,
                        "Erro"
                    ):
        
                        print()
                        print(
                            "GATEWAY TIME-OUT / "
                            "FAVOR VERIFICAR"
                        )
                        time.sleep(180)
        
                        break
        
        
                    print("\rAGUARDANDO VALIDAR IMPORTAÇÃO...",end="",flush=True)
        
                    time.sleep(1)
        
        
                time.sleep(2)
        
                # ========================================================
                # FECHAR AVISO
                # ========================================================
        
                seletor = (
                    "button"
                    "[ng-click='onSuccess()']"
                )
        
                sucesso = clicar_campo(
                    driver,
                    seletor
                )
        
                if not sucesso:
        
                    return {
                        "sucesso": False,
                        "mensagem": (
                            "Não foi possível fechar "
                            "o aviso da validação."
                        )
                    }
        
                time.sleep(2)
        
            else:
            
                print("Nenhuma divergência encontrada.")
                print(
                    "FAVOR REALIZAR MANUALMENTE "
                    "O RESTANTE DO PROCEDIMENTO"
                    )
                
                
            seletor = (
                "button"
                "[ng-click='ctrl.changeFace()']"
            )
            
            clicar_campo(driver, seletor)
    

        # ============================================================
        # PROXIMO PASSO ABRIR TELA DE CENTRAL DE VENDAS
        # ============================================================

        seletor = (
            "button"
            "[ng-click='ctrl.abrirDocumento()']"
        )
        
        clicar_campo(driver, seletor)
        time.sleep(1.5)
        
        mudar_aba(driver,"Sankhya Om")
        time.sleep(1)
        mudar_aba(driver,"Sankhya Om")
        
        time.sleep(20)
        
        seletor = "td > div.Taskbar-icon.icon-user-photo"
        clicar_campo(driver, seletor)
        
        seletor = (
            "//td[@class='gwt-MenuItem' "
            "and @role='menuitem' "
            "and @colspan='2']"
            "[.//span[@class='gwt-MenuItem-icon icon-photo-stack']]"
            "[.//span[normalize-space()='Layout da tela']]"
        )
        
        clicar_campo(driver, seletor)
        
        
        seletor = "input[type='radio'][name='radioHTML5'][value='on']"
        
        clicar_campo(driver,seletor)
        
        seletor = "//button[@type='button' and normalize-space()='OK']"
        
        clicar_campo(driver, seletor)
        
        time.sleep(5)
        
        mudar_aba(driver, "Central de Vendas")
        
        time.sleep(10)
        
        # PREENCHER 202160200
        
        seletor = (
            "//label[.//span[@title='Natureza']]"
            "/following::input[1]"
        )
        
        clicar_campo(driver, seletor)
        digitar_no_foco(driver, "202160200")
        time.sleep(2)
        # DATA DE NEGOCIAÇÃO
        
        seletor = (
            "//label[.//span[@title='Dt. Neg.']]"
            "/following::input[1]"
        )
        
        clicar_campo(driver, seletor)
        digitar_no_foco(driver, data_lançamento)
        time.sleep(2)
        
        
        # # DATA ENTRADA E SAÍDA
        
        seletor = (
            "//label[.//span[@title='Dt. Entrada/Saída']]"
            "/following::input[1]"
        )
        
        clicar_campo(driver, seletor)
        digitar_no_foco(driver, data_lançamento)
        time.sleep(2)
        
        # # DATA DE FATURAMENTO
        
        seletor = (
            "//label[.//span[@title='Dt. do Faturamento']]"
            "/following::input[1]"
        )
        
        clicar_campo(driver, seletor)
        digitar_no_foco(driver,data_lançamento)
        time.sleep(2)
        
        
        ActionChains(driver).send_keys(Keys.F7).perform()
        
        time.sleep(8)
        
        ######################### PREENCHENDO O CODIGO DO MOTIVO
        
        time.sleep(3)
        
        seletor_tabela = (
            "div[role='row'][row-index='0'] "
            "div[role='gridcell'][col-id='AD_CODMOT']"
        )
        
        campos = driver.find_elements(
            By.CSS_SELECTOR,
            seletor_tabela
        )
        
        print(f"Quantidade de linhas para Cód. Motivo: {len(campos)}")
        
        quantidade_linhas = len(campos)
        
        clicar_campo(driver, seletor_tabela)
        
        
        for campo in range(quantidade_linhas):
            digitar_no_foco(driver, motivo)
            time.sleep(2)
            ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
            time.sleep(8)
        
        fechar_aba(driver, "Central de Vendas")
        
        
        
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


