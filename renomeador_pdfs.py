import os
import re
import csv
import shutil
import unicodedata
from pathlib import Path
from datetime import datetime
from tkinter import Tk, Label, Text, END, Radiobutton, StringVar
from tkinter import ttk
import sys

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None
try:
    from PyPDF2 import PdfReader
except Exception:
    PdfReader = None

# ---------------- Utilidades ----------------

def get_base_dir():
    """Retorna a pasta base correta, seja no .py ou no .exe PyInstaller"""
    if getattr(sys, 'frozen', False):  # rodando como .exe
        return Path(sys.executable).parent
    else:  # rodando como script .py
        return Path(__file__).parent

def strip_accents(s: str) -> str:
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def only_digits(s: str) -> str:
    return re.sub(r"\D", "", s)

def read_pdf_text(pdf_path: Path) -> str:
    text_parts = []
    if fitz is not None:
        try:
            with fitz.open(pdf_path) as doc:
                for page in doc:
                    text_parts.append(page.get_text("text"))
        except Exception:
            pass
    text = "\n".join(text_parts).strip()
    if not text and PdfReader is not None:
        try:
            reader = PdfReader(str(pdf_path))
            text = "\n".join([p.extract_text() or "" for p in reader.pages])
        except Exception:
            pass
    return text

def validate_cpf(cpf: str) -> bool:
    cpf = only_digits(cpf)
    if len(cpf) != 11 or cpf == cpf[0]*11:
        return False
    soma = sum(int(cpf[i])*(10-i) for i in range(9))
    dv1 = (soma*10)%11
    if dv1==10: dv1=0
    if dv1 != int(cpf[9]): return False
    soma = sum(int(cpf[i])*(11-i) for i in range(10))
    dv2 = (soma*10)%11
    if dv2==10: dv2=0
    return dv2 == int(cpf[10])

# ---------------- Extrações ----------------

CPF_REGEX = re.compile(r"\b\d{3}[\.-]?\d{3}[\.-]?\d{3}[-]?\d{2}\b")
CPF_ANY11_REGEX = re.compile(r"\b\d{11}\b")

ANCHORS_COMPRADOR = [r"COMPRADOR", r"DEVEDOR", r"PARTE ADQUIRENTE", r"PARTE COMPRADORA"]

def extract_cpf_first_buyer(text: str) -> str|None:
    # escopo a partir de palavras-âncora
    t_up = strip_accents(text.upper())
    pos = None
    for a in ANCHORS_COMPRADOR:
        idx = t_up.find(a)
        if idx != -1:
            pos = idx
            break
    scope = text[pos:] if pos is not None else text
    m = CPF_REGEX.search(scope) or CPF_ANY11_REGEX.search(scope)
    if not m:
        return None
    cpf = only_digits(m.group(0))
    if not validate_cpf(cpf):
        return None
    return cpf

def extract_contract_number(text: str) -> str|None:
    t = strip_accents(text.upper())
    m = re.search(r"CONTRATO.*?([\d\.-]{8,})", t)
    if m:
        digits = only_digits(m.group(1))
        if len(digits) >= 13:
            return digits[-13:]
    m2 = re.search(r"\b(\d{13})\b", t)
    if m2:
        return m2.group(1)
    return None

def extract_nome_until_comma(text: str, anchors_regexes: list[str]) -> str|None:
    """
    Procura por qualquer âncora em anchors_regexes (após normalização -> maiúsculas e sem acentos),
    e retorna o trecho subsequente até a PRIMEIRA vírgula, concatenando quebras de linha.
    """
    raw = text
    up = strip_accents(raw.upper())

    # encontre primeira âncora que der match
    start_idx = None
    end_match = None
    for rgx in anchors_regexes:
        m = re.search(rgx, up)
        if m:
            start_idx = m.end()
            end_match = m
            break
    if start_idx is None:
        return None

    tail_up = up[start_idx:]  # texto já normalizado
    # procura a primeira vírgula no tail (na versão normalizada mantemos vírgula)
    comma_pos = tail_up.find(',')
    if comma_pos == -1:
        # fallback: pega a primeira linha não vazia, como antes
        lines = [ln.strip() for ln in tail_up.splitlines() if ln.strip()]
        if not lines:
            return None
        candidate = lines[0]
    else:
        candidate = tail_up[:comma_pos]

    # Limpa caracteres que não sejam letras/espacos/quebras (antes de colar)
    candidate = re.sub(r"[^A-Z \n]", " ", candidate)
    # Une quebras de linha e normaliza múltiplos espaços
    candidate = re.sub(r"\s+", " ", candidate).strip()

    # Evita retorno vazio
    return candidate if candidate else None

def extract_oficio_num(text: str) -> str:
    """
    Se detectar 6º ofício, retorna '6'. Caso contrário, assume '5' como fallback.
    Aceita variações como '6º', '6o', '6 O', 'OFICIO' sem acento, etc.
    """
    txt = strip_accents(text.upper())
    if re.search(r"\b6[ºO]?\s*OFICIO\b", txt):
        return "6"
    return "5"

# ---------------- Casos ----------------

def rename_contratos(pdf, outdir):
    text = read_pdf_text(pdf)
    if not text:
        return (pdf.name, "", "ERRO: PDF sem texto")
    cpf = extract_cpf_first_buyer(text)
    contrato = extract_contract_number(text)
    if not cpf:
        return (pdf.name, "", "ERRO: CPF não encontrado")
    if not contrato:
        return (pdf.name, "", "ERRO: Contrato não encontrado")
    novo = f"{cpf}_{contrato}.pdf"
    dest = outdir/novo
    i=1
    while dest.exists():
        dest = outdir/f"{cpf}_{contrato}_{i}.pdf"; i+=1
    shutil.copy2(pdf, dest)
    return (pdf.name, dest.name, "OK")

def rename_certidoes_2(pdf, outdir):
    text = read_pdf_text(pdf)
    if not text:
        return (pdf.name, "", "ERRO: PDF sem texto")
    # Mantém âncora original para -2
    nome = extract_nome_until_comma(text, [
        r"COM\s+REFERENCIA\s+AO\s+NOME\s+DE"
    ])
    if not nome:
        return (pdf.name, "", "ERRO: Nome não encontrado")
    dest = outdir/f"{nome}-2.pdf"
    i=1
    while dest.exists():
        dest = outdir/f"{nome}-2.{i}.pdf"; i+=1
    shutil.copy2(pdf, dest)
    return (pdf.name, dest.name, "OK")

def rename_certidoes_5_6(pdf, outdir):
    text = read_pdf_text(pdf)
    if not text:
        return (pdf.name, "", "ERRO: PDF sem texto")

    # Âncoras aceitas:
    # - "NADA CONSTA EM NOME DE"
    # - "EM NOME DE"
    nome = extract_nome_until_comma(text, [
        r"NADA\s+CONSTA\s+EM\s+NOME\s+DE",
        r"EM\s+NOME\s+DE",
    ])
    oficio = extract_oficio_num(text)  # '6' se achar 6º, senão '5'

    if not nome:
        return (pdf.name, "", "ERRO: Nome não encontrado")

    dest = outdir/f"{nome}-{oficio}.pdf"
    i=1
    while dest.exists():
        dest = outdir/f"{nome}-{oficio}.{i}.pdf"; i+=1
    shutil.copy2(pdf, dest)
    return (pdf.name, dest.name, "OK")

# ---------------- Interface ----------------
class App:
    def __init__(self, master):
        self.master = master
        master.title("Renomeador de PDFs")
        master.geometry("820x540")
        master.configure(bg="#f0f6ff")

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TButton", background="#007BFF", foreground="white", padding=6, font=("Segoe UI", 10, "bold"))
        style.map("TButton", background=[("active", "#0056b3")])

        self.mode = StringVar(value="contratos")

        Label(master, text="Renomeador de PDFs", bg="#f0f6ff", font=("Segoe UI", 16, "bold"), fg="#007BFF").pack(pady=10)

        Label(master, text="Escolha o tipo de documentos:", bg="#f0f6ff", font=("Segoe UI", 12)).pack(pady=5)
        Radiobutton(master, text="Contratos", variable=self.mode, value="contratos", bg="#f0f6ff").pack(anchor='w', padx=40)
        Radiobutton(master, text="Certidões -2", variable=self.mode, value="certidoes_2", bg="#f0f6ff").pack(anchor='w', padx=40)
        Radiobutton(master, text="Certidões 5/6", variable=self.mode, value="certidoes_5_6", bg="#f0f6ff").pack(anchor='w', padx=40)

        ttk.Button(master, text="Executar", command=self.run).pack(pady=15)

        Label(master, text="Log:", bg="#f0f6ff").pack()
        self.log = Text(master, height=18)
        self.log.pack(fill='both', expand=True, padx=10, pady=6)

        # Criação das pastas na inicialização
        self.ensure_folders()

    def ensure_folders(self):
        base = get_base_dir()
        for p in ["contratos", "certidoes_2", "certidoes_5_6"]:
            (base/"entrada"/p).mkdir(parents=True, exist_ok=True)
            (base/"saida"/p).mkdir(parents=True, exist_ok=True)

    def run(self):
        base = get_base_dir()
        entrada = base/"entrada"/self.mode.get()
        saida = base/"saida"/self.mode.get()

        fn = {
            "contratos": rename_contratos,
            "certidoes_2": rename_certidoes_2,
            "certidoes_5_6": rename_certidoes_5_6
        }[self.mode.get()]

        arquivos = sorted(entrada.glob("*.pdf"))
        if not arquivos:
            self.log.insert(END, f"⚠ A pasta '{entrada}' está vazia. Nenhum PDF para processar.\n")
            return

        ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        logcsv = saida/f"renomeacao_log_{ts}.csv"
        csvf = open(logcsv,'w',newline='',encoding='utf-8')
        writer = csv.writer(csvf, delimiter=';')
        writer.writerow(["Arquivo Original","Nome Novo","Status"])

        ok = 0
        for pdf in arquivos:
            try:
                orig, novo, status = fn(pdf, saida)
            except Exception as e:
                orig, novo, status = (pdf.name, "", f"ERRO: {e}")
            writer.writerow([orig, novo, status])
            if status=="OK":
                self.log.insert(END,f"✔ {orig} → {novo}\n"); ok+=1
            else:
                self.log.insert(END,f"✖ {orig} → {status}\n")
        csvf.close()
        self.log.insert(END,f"\nConcluído. {ok} arquivos renomeados. Log salvo em {logcsv}\n")

if __name__=="__main__":
    root = Tk()
    app = App(root)
    root.mainloop()
