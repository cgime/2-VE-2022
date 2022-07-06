from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from typing import Dict, List
from prettytable import PrettyTable
driver = webdriver.Chrome(ChromeDriverManager().install())
arquivo_nome_numero_CGs = "automacao_notas\CGs.txt"
arquivo_final_classificao = "automacao_notas\classificacaoCGs.txt"

# -------------------- PARTE 1: FAZER O LOGIN --------------------
driver.get("http://granito.ime.eb.br/Granito/SisGradWeb/login.asp")
txt = driver.find_element(By.NAME, "txtUser")
txt.clear()
txt.send_keys("20044")
senha = driver.find_element(By.NAME, "txtSenha")
senha.clear()
senha.send_keys("20044")
botao = driver.find_element(By.XPATH, "/html/body/form/center/table/tbody/tr/td[5]/input")
botao.click()

# ------------------ PARTE 2: IR PARA O LOCAL ONDE ACESSA A NOTA --------------
driver.get("http://granito.ime.eb.br/Granito/SisGradWeb/BolGrausSolicit.asp")
data = driver.find_element(By.ID, "text3")
driver.execute_script("arguments[0].value = arguments[1]", data, "1") # Mudar periodo para o primeiro

# ------------------ PARTE 3: obter os nomes/numeros --------------------------
cgs: Dict[str, Dict] = {}
indice: int = 0
nome: str = ""
numero: str = ""
with open(arquivo_nome_numero_CGs, "r") as arq:
    for linha in arq:
        if indice % 2 == 0:
            numero = linha[:len(linha)-1]
        else:
            nome = linha[:len(linha)-1]
        cgs[numero] = {
            'nome': nome,
            'NCF': 0
        }
        indice += 1

# ----------------------- PARTE 4: OBTER O NFC --------------------------
for cg in cgs.keys():
    ncf: List[int] = []
    txt = driver.find_element(By.ID, "text1")
    driver.execute_script("arguments[0].value = arguments[1]", txt, cg)
    driver.find_element(By.ID, "submit1").click()
    driver.switch_to.window(driver.window_handles[1])
    nota = driver.find_element(By.XPATH, "//*[@id='form1']/font/table[4]/tbody/tr/td[2]").text
    # formatar a nota
    ncf_final_aux = []
    for ch in nota:
        if ch.isnumeric():
            ncf_final_aux.append(ch)
        if ch == ',':
            ncf_final_aux.append('.')
    ncf_final = float("".join(ncf_final_aux))
    
    # adicionar a nota ao dicionario
    cgs[cg]['NCF'] = ncf_final
    ncf_final_aux.clear()
    driver.close()
    driver.switch_to.window(driver.window_handles[0])

# --------------------- PARTE 6: ORGANIZAR POR NOTAS ---------------------
cgs = dict(sorted(cgs.items(), key=lambda x: x[1]['NCF'], reverse=True))

# --------------------- PARTE 5: FORMATAR NA TABELA ---------------------
tabela = PrettyTable()
tabela.border = False
tabela.preserve_internal_border = True
tabela.field_names = ["POSICAO", "NUMERO", "NOME", "NCF"]

posicao: str = ""
numero: str = ""
nome: str = ""
ncf: str = ""
for pos, cg in enumerate(cgs.keys()):
    numero = cg
    if pos < 9:
        posicao = "0" + str(pos + 1)
    else:
        posicao = str(pos + 1)
    nome = cgs[cg]['nome']
    ncf = cgs[cg]['NCF']
    tabela.add_row([posicao, numero, nome, ncf])
    
resultado: str = tabela.get_string()
with open(arquivo_final_classificao, "w") as arq:
    arq.write(resultado)






