import requests
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime
import os
import warnings

warnings.filterwarnings("ignore", message="Unverified HTTPS request")

# =======================
# CONFIGURAÇÕES
# =======================
API_URL_BASE = "https://MAN-prd.jemsms.corp.jabil.org/"
USER = r"jabil\svchua_jesmapistg"
PASSWORD = "qKzla3oBDA51Ecq=+B2_z"

OUTPUT_DIR = r"\\manfile01\!General\Samsung\TE\106 - MES APPLICATION\SAMPALLET REPORT"
os.makedirs(OUTPUT_DIR, exist_ok=True)

SITE_CODE = "MAN"
CUSTOMER_ID = 2

_cached_token = None

# =======================
# AUTENTICAÇÃO API
# =======================
def get_token():
    global _cached_token
    if _cached_token:
        return _cached_token
    url = f"{API_URL_BASE}api-external-api/api/user/adsignin"
    form_data = {"name": USER, "password": PASSWORD}
    try:
        r = requests.post(url, data=form_data, verify=False, timeout=15)
        r.raise_for_status()
        _cached_token = r.text.strip()
        print("✅ Token obtido com sucesso")
        return _cached_token
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao autenticar: {e}")
        return None

def build_session():
    token = get_token()
    if not token:
        return None
    session = requests.Session()
    session.verify = False
    if "=" in token:
        name, value = token.split("=", 1)
        session.cookies.set(name.strip(), value.strip(";"))
    else:
        session.cookies.set("AuthToken", token)
    return session

# =======================
# API REQUEST
# =======================
def get_container_info(session, serial):
    endpoint = f"{API_URL_BASE}api-external-api/api/containers/containerhierarchy/contentsByWip/external"
    params = {
        "SiteCode": SITE_CODE,
        "WipSerialNumber": serial,
        "CustomerId": CUSTOMER_ID
    }
    try:
        r = session.get(endpoint, params=params, timeout=20)
        if r.status_code != 200:
            raise Exception(f"HTTP {r.status_code}: {r.text[:200]}")
        return r.json()
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao consultar API:\n{e}")
        return None

# =======================
# EXPORTAÇÃO / LOG
# =======================
def save_to_log(data, filename_base):
    log_file = os.path.join(OUTPUT_DIR, f"{filename_base}_log.txt")
    try:
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {data}\n\n")
    except Exception as e:
        print("Erro ao salvar log:", e)

def export_to_excel(data, filename_base):
    excel_file = os.path.join(OUTPUT_DIR, f"{filename_base}_export.xlsx")
    try:
        if not data:
            messagebox.showwarning("Aviso", "Sem dados para exportar.")
            return

        parent = {
            "ContainerNumber": data.get("ContainerNumber", "N/A"),
            "ContainerStatus": data.get("ContainerStatus", "N/A"),
            "ContainerUsageType": data.get("ContainerUsageType", "N/A"),
            "ContainerCloseDate": data.get("ContainerCloseDate", "N/A"),
            "ChildContainersCount": data.get("ChildContainersCount", 0),
        }

        rows = []
        for child in data.get("ChildContainers", []):
            det = child.get("ContainerDetails", {})
            for wip in det.get("WIPSerialNumbers", []):
                rows.append({
                    **parent,
                    "ChildContainerNumber": child.get("ContainerNumber", "N/A"),
                    "Material": det.get("Material", "N/A"),
                    "AssemblyNumber": det.get("AssemblyNumber", "N/A"),
                    "AssemblyRevision": det.get("AssemblyRevision", "N/A"),
                    "AssemblyVersion": det.get("AssemblyVersion", "N/A"),
                    "PackedDate": det.get("ContainerPackedDate", "N/A"),
                    "SerialNumber": wip.get("SerialNumber", "N/A"),
                })

        df = pd.DataFrame(rows)
        if df.empty:
            messagebox.showwarning("Aviso", "Nenhum dado para exportar.")
            return

        df.to_excel(excel_file, index=False, sheet_name="ContainerHierarchy")
        messagebox.showinfo("Exportar Excel", f"✅ Dados exportados:\n{excel_file}")

    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao exportar Excel:\n{e}")

# =======================
# TKINTER UI
# =======================
class ContainerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("📦 Jabil MES - Container Hierarchy Extractor")
        self.geometry("650x450")
        self.resizable(False, False)
        self.session = build_session()
        if not self.session:
            self.destroy()
            return
        self.create_widgets()

    def create_widgets(self):
        ttk.Label(self, text="Digite o Serial WIP:").pack(pady=5)
        self.serial_entry = ttk.Entry(self, width=40)
        self.serial_entry.pack()
        self.serial_entry.bind("<Return>", lambda e: self.on_extract())

        ttk.Label(self, text="Nome base do arquivo:").pack(pady=5)
        self.filename_entry = ttk.Entry(self, width=40)
        self.filename_entry.pack()

        ttk.Button(self, text="🔍 Extrair", command=self.on_extract).pack(pady=10)
        ttk.Button(self, text="📤 Exportar Excel/Log", command=self.export_last).pack(pady=5)

        self.result_box = tk.Text(self, width=80, height=15)
        self.result_box.pack(pady=10)

        self.last_data = None

    def on_extract(self):
        serial = self.serial_entry.get().strip()
        if not serial:
            messagebox.showwarning("Aviso", "Digite um número de série.")
            return

        filename_base = self.filename_entry.get().strip() or "container_hierarchy"

        self.result_box.delete("1.0", tk.END)
        self.result_box.insert(tk.END, f"Consultando serial {serial}...\n")

        data = get_container_info(self.session, serial)
        if not data:
            self.result_box.insert(tk.END, "❌ Nenhum dado retornado ou erro na API.\n")
            return

        self.last_data = data
        save_to_log(data, filename_base)
        self.result_box.insert(tk.END, f"✅ Container: {data.get('ContainerNumber')}\n")
        self.result_box.insert(tk.END, f"Status: {data.get('ContainerStatus')}\n")
        self.result_box.insert(tk.END, f"ChildContainers: {len(data.get('ChildContainers', []))}\n")
        self.result_box.insert(tk.END, "\nExemplo do primeiro ChildContainer:\n")
        if data.get("ChildContainers"):
            first = data["ChildContainers"][0]
            self.result_box.insert(tk.END, f" - {first.get('ContainerNumber')}\n")

    def export_last(self):
        if not self.last_data:
            messagebox.showinfo("Aviso", "Nenhum dado carregado ainda.")
            return
        filename_base = self.filename_entry.get().strip() or "container_hierarchy"
        export_to_excel(self.last_data, filename_base)
        save_to_log(self.last_data, filename_base)

# =======================
# EXECUÇÃO
# =======================
if __name__ == "__main__":
    app = ContainerApp()
    app.mainloop()
