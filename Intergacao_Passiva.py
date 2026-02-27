import ast
from pathlib import Path
import pymysql
import os
import time
import re
import signal
import sys

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException

_sigint_count = 0

def _sigint_handler(sig, frame):
    global _sigint_count
    _sigint_count += 1
    print(f"\nSIGINT recebido ({_sigint_count}).")
    if _sigint_count >= 2:
        raise KeyboardInterrupt

signal.signal(signal.SIGINT, _sigint_handler)

def log(msg: str):
    print(msg)

CRED_PATH = r"\\fs01\ITAPEVA ATIVAS\DADOS\SA_Credencials.txt"

ALPHERATZ_URL = "https://alpheratz.exemplo.com.br/Alpheratz/intranet/default.aspx"
ALPHERATZ_USER = "<SEU_USUARIO_ALPHERATZ>"
ALPHERATZ_PASS = "<SUA_SENHA_ALPHERATZ>"

CHROME_PROFILE_DIR = r"C:\Temp\selenium_alpheratz_profile"

QUERY = """
SELECT 
    tra.data_hora_lan,
    tra.update_usuario,
    tra.texto,
    tra.evento,
    cad.numero_integracao,
    cad.pj,
    eve.descricao
FROM tramitacao tra
JOIN cad_processo cad ON cad.pj = tra.id_processo
JOIN tab_evento eve ON eve.sigla = tra.evento
WHERE tra.evento = 'itp45'
  AND tra.data_hora_lan >= CURDATE()
  AND tra.data_hora_lan < CURDATE() + INTERVAL 1 DAY
"""

def load_vars_from_txt(txt_path: str) -> dict:
    text = Path(txt_path).read_text(encoding="utf-8", errors="ignore")
    tree = ast.parse(text, filename=str(txt_path))
    values = {}
    for node in tree.body:
        if isinstance(node, ast.Assign) and len(node.targets) == 1 and isinstance(node.targets[0], ast.Name):
            name = node.targets[0].id
            try:
                values[name] = ast.literal_eval(node.value)
            except Exception:
                pass
    return values

def load_cpjwcs_config(txt_path: str) -> dict:
    v = load_vars_from_txt(txt_path)
    required = ["CPJWCS_HOST", "CPJWCS_USER", "CPJWCS_PASS", "CPJWCS_DB", "CPJWCS_PORT"]
    missing = [k for k in required if k not in v]
    if missing:
        raise RuntimeError(f"Credenciais CPJWCS incompletas no arquivo. Faltando: {missing}")
    return {
        "host": v["CPJWCS_HOST"],
        "user": v["CPJWCS_USER"],
        "password": v["CPJWCS_PASS"],
        "database": v["CPJWCS_DB"],
        "port": int(v["CPJWCS_PORT"]),
        "charset": "utf8mb4",
        "cursorclass": pymysql.cursors.DictCursor,
        "connect_timeout": 10,
        "read_timeout": 60,
        "write_timeout": 60,
    }

def connect_cpjwcs():
    cfg = load_cpjwcs_config(CRED_PATH)
    return pymysql.connect(**cfg)

def run_query(conn):
    with conn.cursor() as cursor:
        cursor.execute(QUERY)
        return cursor.fetchall()

def normalizar_frase(s: str) -> str:
    if not s:
        return ""
    s = str(s).lower()
    s = s.replace("\ufeff", "").replace("\ufffd", " ").replace("�", " ")
    s = re.sub(r"[^\w\s]+", " ", s, flags=re.UNICODE)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def tokens_relevantes(s: str) -> set:
    n = normalizar_frase(s)
    toks = [t for t in n.split() if len(t) >= 3]
    return set(toks)

def similaridade_token_set(a: str, b: str) -> float:
    A = tokens_relevantes(a)
    B = tokens_relevantes(b)
    if not A or not B:
        return 0.0
    inter = len(A & B)
    base = min(len(A), len(B))
    return inter / base if base else 0.0

def _wait_dom_ready(driver, timeout=25):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") in ("interactive", "complete")
    )

def _first_visible(elements):
    for el in elements:
        try:
            if el.is_displayed() and el.is_enabled():
                return el
        except Exception:
            continue
    return None

def _safe_click(driver, el, pause=0.25):
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center', inline:'center'});", el)
        time.sleep(pause)
    except Exception:
        pass
    try:
        el.click()
        time.sleep(pause)
        return
    except Exception:
        pass
    try:
        ActionChains(driver).move_to_element(el).pause(pause).click(el).perform()
        time.sleep(pause)
        return
    except Exception:
        pass
    try:
        driver.execute_script("arguments[0].click();", el)
        time.sleep(pause)
    except Exception:
        pass

def _clear_and_type(el, value: str):
    if value is None:
        value = ""
    value = str(value).strip()
    try:
        el.click()
    except Exception:
        pass
    try:
        el.send_keys(Keys.CONTROL, "a")
        el.send_keys(Keys.DELETE)
    except Exception:
        pass
    if value:
        el.send_keys(value)

def _hit_escape(driver, pause=0.08):
    try:
        driver.switch_to.active_element.send_keys(Keys.ESCAPE)
    except Exception:
        try:
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        except Exception:
            pass
    time.sleep(pause)

def clicar_salvar(driver, timeout=12) -> bool:
    try:
        wait = WebDriverWait(driver, timeout)
        btn = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_content_bt_step_save")))
        _safe_click(driver, btn, pause=0.35)
        _wait_dom_ready(driver, timeout=25)
        return True
    except Exception:
        return False

def preencher_tela_por_texto(driver, texto_db: str) -> int:
    FORCE_SELECT_WHEN_EMPTY = {"cumprimento integral"}

    def _norm(s: str) -> str:
        if not s:
            return ""
        s = str(s).replace("\ufeff", "").replace("\ufffd", " ").replace("�", " ")
        s = s.strip().lower()
        s = re.sub(r"\s+", " ", s)
        s = s.rstrip(":").strip()
        return s

    def _parse_kv(txt: str) -> dict:
        if not txt:
            return {}
        t = str(txt).replace("\ufeff", "").replace("\ufffd", " ").replace("�", " ")
        kv = {}
        for ln in t.splitlines():
            ln = ln.strip()
            if not ln or ":" not in ln:
                continue
            k, v = ln.split(":", 1)
            k = (k or "").strip().rstrip(":").strip()
            v = (v or "").strip()
            if k:
                kv[_norm(k)] = v
        return kv

    def _select_by_text_flex(select_el, desired: str) -> bool:
        desired = (desired or "").strip()
        if not desired:
            return False
        dn = _norm(desired)
        sel = Select(select_el)
        for opt in sel.options:
            ot = (opt.text or "").strip()
            if ot and _norm(ot) == dn:
                sel.select_by_visible_text(ot)
                return True
        for opt in sel.options:
            ot = (opt.text or "").strip()
            if not ot:
                continue
            on = _norm(ot)
            if dn and (dn in on or on in dn):
                sel.select_by_visible_text(ot)
                return True
        return False

    def _is_yes_no(value: str):
        v = _norm(value)
        yes = v in ("sim", "s", "yes", "y", "1", "true", "verdadeiro")
        no = v in ("nao", "não", "n", "no", "0", "false", "falso")
        return yes, no

    def _multa_valor_para_bool(v: str) -> bool:
        if v is None:
            return False
        s = str(v).strip()
        if not s:
            return False
        s2 = s.replace(" ", "").replace(".", "").replace(",", ".")
        try:
            return float(s2) > 0
        except Exception:
            return bool(re.search(r"\d", s))

    def _click_radio_sim_nao(campo_td, desired_value: str) -> bool:
        yes, no = _is_yes_no(desired_value)
        if not (yes or no):
            return False
        try:
            if yes:
                rb = campo_td.find_element(By.CSS_SELECTOR, "input[type='radio']#ctl00_content_rbyes4")
            else:
                rb = campo_td.find_element(By.CSS_SELECTOR, "input[type='radio']#ctl00_content_rbno4")
            _safe_click(driver, rb, pause=0.20)
            return True
        except Exception:
            pass
        wanted = "sim" if yes else "nao"
        try:
            labels = campo_td.find_elements(By.TAG_NAME, "label")
            for lb in labels:
                try:
                    txt = _norm(lb.text)
                    if txt == wanted or (txt == "não" and wanted == "nao"):
                        fid = lb.get_attribute("for")
                        if fid:
                            rb = campo_td.find_element(By.ID, fid)
                            _safe_click(driver, rb, pause=0.20)
                            return True
                except Exception:
                    continue
        except Exception:
            pass
        try:
            radios = campo_td.find_elements(By.CSS_SELECTOR, "input[type='radio']")
            radios = [r for r in radios if r.is_displayed() and r.is_enabled()]
            if len(radios) >= 2:
                _safe_click(driver, radios[0] if yes else radios[1], pause=0.20)
                return True
        except Exception:
            pass
        return False

    def _select_default_if_empty(select_el) -> bool:
        sel = Select(select_el)
        try:
            sel.select_by_value("-1")
            return True
        except Exception:
            pass
        try:
            if sel.options:
                sel.select_by_index(0)
                return True
        except Exception:
            pass
        return False

    kv = _parse_kv(texto_db)
    if not kv:
        return 0

    multa_key = _norm("Multa Diária")
    valor_multa_key = _norm("Valor da Multa Diária")
    if multa_key not in kv and valor_multa_key in kv:
        kv[multa_key] = "SIM" if _multa_valor_para_bool(kv.get(valor_multa_key)) else "NÃO"

    wait = WebDriverWait(driver, 25)
    wait.until(EC.presence_of_element_located((By.XPATH, "//tbody")))
    trs = driver.find_elements(By.CSS_SELECTOR, "tbody > tr")
    preenchidos = 0

    for tr in trs:
        try:
            tds = tr.find_elements(By.TAG_NAME, "td")
            if len(tds) < 3:
                continue
            box_html = (tds[0].text or "").strip()
            nbox = _norm(box_html)
            if not nbox:
                continue
            if nbox not in kv:
                continue
            value = kv.get(nbox, "")
            value = "" if value is None else str(value).strip()
            campo_td = tds[2]
            try:
                any_radio = _first_visible(campo_td.find_elements(By.CSS_SELECTOR, "input[type='radio']"))
                if any_radio:
                    if value == "":
                        continue
                    ok = _click_radio_sim_nao(campo_td, value)
                    _hit_escape(driver)
                    if ok:
                        preenchidos += 1
                    continue
            except StaleElementReferenceException:
                continue
            except Exception:
                pass
            try:
                sel_el = _first_visible(campo_td.find_elements(By.TAG_NAME, "select"))
                if sel_el:
                    _safe_click(driver, sel_el, pause=0.20)
                    if value == "":
                        if nbox in FORCE_SELECT_WHEN_EMPTY:
                            ok = _select_default_if_empty(sel_el)
                            _hit_escape(driver)
                            if ok:
                                preenchidos += 1
                            continue
                        else:
                            continue
                    ok = _select_by_text_flex(sel_el, value)
                    _hit_escape(driver)
                    if ok:
                        preenchidos += 1
                    continue
            except StaleElementReferenceException:
                continue
            except Exception:
                pass
            if value == "":
                continue
            try:
                ta = _first_visible(campo_td.find_elements(By.TAG_NAME, "textarea"))
                if ta:
                    _safe_click(driver, ta, pause=0.20)
                    _clear_and_type(ta, value)
                    _hit_escape(driver)
                    preenchidos += 1
                    continue
            except StaleElementReferenceException:
                continue
            except Exception:
                pass
            try:
                inp = _first_visible(campo_td.find_elements(By.CSS_SELECTOR, "input"))
                if inp:
                    itype = (inp.get_attribute("type") or "").strip().lower()
                    if itype in ("file", "submit", "button", "image"):
                        continue
                    _safe_click(driver, inp, pause=0.20)
                    _clear_and_type(inp, value)
                    _hit_escape(driver)
                    preenchidos += 1
                    continue
            except StaleElementReferenceException:
                continue
            except Exception:
                pass
        except StaleElementReferenceException:
            continue
        except Exception:
            continue

    return preenchidos

def clicar_proxima_etapa_se_existir(driver, descricao_query: str, timeout=3) -> bool:
    SIMILARIDADE_MINIMA = 0.78
    wait = WebDriverWait(driver, timeout)
    descricao_query = (descricao_query or "").strip()
    if not descricao_query:
        return False
    try:
        td = wait.until(EC.presence_of_element_located((By.ID, "ctl00_content_td_nextstep")))
    except Exception:
        return False
    try:
        btns = td.find_elements(By.CSS_SELECTOR, "input.bt-go")
    except Exception:
        btns = []
    btn = _first_visible(btns)
    if not btn:
        return False
    texto_btn = (btn.get_attribute("value") or btn.text or "").strip()
    score = similaridade_token_set(descricao_query, texto_btn)
    if score >= SIMILARIDADE_MINIMA:
        _safe_click(driver, btn, pause=0.35)
        _wait_dom_ready(driver, timeout=25)
        return True
    btn_tokens = tokens_relevantes(texto_btn)
    desc_tokens = tokens_relevantes(descricao_query)
    chaves = {"central", "liminares"}
    if chaves <= btn_tokens and ("central" in desc_tokens or "liminares" in desc_tokens):
        _safe_click(driver, btn, pause=0.35)
        _wait_dom_ready(driver, timeout=25)
        return True
    return False

def build_driver():
    os.makedirs(CHROME_PROFILE_DIR, exist_ok=True)
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument(f"--user-data-dir={CHROME_PROFILE_DIR}")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
    return webdriver.Chrome(options=chrome_options)

def alpheratz_open_and_login(driver, username: str, password: str):
    driver.get(ALPHERATZ_URL)
    _wait_dom_ready(driver, timeout=25)
    try:
        WebDriverWait(driver, 8).until(
            EC.presence_of_element_located(
                (By.XPATH, "//input[translate(@type,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')='password']")
            )
        )
    except Exception:
        return
    user_el = None
    try:
        prev_inputs = driver.find_elements(
            By.XPATH,
            "//input[translate(@type,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')='password']"
            "/preceding::input[not(@type='hidden')]"
        )
        user_el = _first_visible(prev_inputs[::-1])
    except Exception:
        user_el = None
    if not user_el:
        user_el = _first_visible(
            driver.find_elements(
                By.XPATH,
                "//input[not(@type='hidden') and ("
                "translate(@type,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')='text' or "
                "translate(@type,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')='email'"
                ")]"
            )
        )
    if not user_el:
        return
    _safe_click(driver, user_el, pause=0.20)
    user_el.send_keys(Keys.CONTROL, "a")
    user_el.send_keys(Keys.DELETE)
    user_el.send_keys(username)
    user_el.send_keys(Keys.TAB)
    active = driver.switch_to.active_element
    active.send_keys(password)
    active.send_keys(Keys.TAB)
    active.send_keys(Keys.ENTER)
    _wait_dom_ready(driver, timeout=25)

def navegar_para_legal(driver):
    wait = WebDriverWait(driver, 30)
    try:
        driver.switch_to.default_content()
    except Exception:
        pass
    _wait_dom_ready(driver, timeout=25)
    legal_candidates = driver.find_elements(By.CSS_SELECTOR, "a#menuitemsTEST[href*='legal.aspx']")
    if not legal_candidates:
        legal_candidates = driver.find_elements(By.ID, "menuitemsTEST")
    legal_btn = _first_visible(legal_candidates)
    if not legal_btn:
        wait.until(EC.presence_of_all_elements_located((By.ID, "menuitemsTEST")))
        legal_candidates = driver.find_elements(By.ID, "menuitemsTEST")
        legal_btn = _first_visible(legal_candidates)
    if not legal_btn:
        raise TimeoutException("Botão Legal não encontrado.")
    _safe_click(driver, legal_btn, pause=0.25)
    try:
        wait.until(EC.url_contains("ft5/legal.aspx"))
    except Exception:
        pass
    return wait.until(EC.visibility_of_element_located((By.ID, "ctl00_content_txt_search")))

def pesquisar_numero_integracao(driver, search_input, numero_integracao: str):
    numero_integracao = (numero_integracao or "").strip()
    if not numero_integracao:
        raise ValueError("numero_integracao vazio.")
    _safe_click(driver, search_input, pause=0.20)
    try:
        search_input.send_keys(Keys.CONTROL, "a")
        search_input.send_keys(Keys.DELETE)
    except Exception:
        pass
    search_input.send_keys(numero_integracao)
    search_input.send_keys(Keys.ENTER)

def clicar_litigation_e_workflow(driver):
    wait = WebDriverWait(driver, 30)
    litigation_span = wait.until(
        EC.element_to_be_clickable((By.ID, "ctl00_content_rpt_OmniIndex_ctl01_lbl_OmniIndex_LitigationID"))
    )
    _safe_click(driver, litigation_span, pause=0.30)
    workflow_btn = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_content_bt_workflow")))
    _safe_click(driver, workflow_btn, pause=0.35)

def open_login_and_go_legal_with_retry(max_retries: int = 1):
    last_err = None
    for attempt in range(max_retries + 1):
        driver = None
        try:
            driver = build_driver()
            alpheratz_open_and_login(driver, ALPHERATZ_USER, ALPHERATZ_PASS)
            WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located((By.ID, "menuitemsTEST")))
            search_input = navegar_para_legal(driver)
            return driver, search_input
        except Exception as e:
            last_err = e
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass
            if attempt < max_retries:
                continue
            raise last_err

if __name__ == "__main__":
    conn = None
    driver = None
    itens = []
    try:
        try:
            conn = connect_cpjwcs()
            log("Banco: CONECTADO")
            rows = run_query(conn)
            for r in rows:
                ni = (r.get("numero_integracao") or "").strip()
                texto_db = (r.get("texto") or "")
                descricao = (r.get("descricao") or "").strip()
                if not ni:
                    continue
                itens.append({"numero_integracao": ni, "texto": texto_db, "descricao": descricao})
            log(f"Contratos encontrados hoje (itp45): {len(itens)}")
        except Exception as e:
            log(f"Banco: FALHOU ({e})")
        finally:
            if conn:
                try:
                    conn.close()
                except Exception:
                    pass
        if not itens:
            log("Nenhum contrato para processar.")
            raise SystemExit(0)
        driver, search_input = open_login_and_go_legal_with_retry(max_retries=1)
        log("Alpheratz: OK")
        for idx, item in enumerate(itens, start=1):
            ni = item["numero_integracao"]
            texto_db = item.get("texto", "") or ""
            descricao = item.get("descricao", "") or ""
            log(f"[{idx}/{len(itens)}] Contrato: {ni}")
            pesquisar_numero_integracao(driver, search_input, ni)
            time.sleep(1.0)
            clicar_litigation_e_workflow(driver)
            time.sleep(0.9)
            clicou = clicar_proxima_etapa_se_existir(driver, descricao_query=descricao, timeout=4)
            if clicou:
                time.sleep(0.8)
            preenchidos = preencher_tela_por_texto(driver, texto_db)
            log(f"Campos preenchidos: {preenchidos}")
            salvou = clicar_salvar(driver, timeout=12)
            if salvou:
                log("Salvar: OK")
            else:
                log("Salvar: não encontrado.")
            time.sleep(0.6)
        try:
            clicar_salvar(driver, timeout=6)
        except Exception:
            pass
    except KeyboardInterrupt:
        log("Interrompido pelo usuário.")
    except Exception as e:
        log(f"Erro geral: {e}")
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
        log("Finalizado.")
        sys.exit(0)