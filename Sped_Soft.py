import os
import sys
import csv
import json
import requests
import re
import pandas as pd
from datetime import datetime
import time
import webbrowser
import smtplib
import urllib.parse
import textwrap
import socket
import getpass
from email.message import EmailMessage
from PySide6.QtCore import QUrl, Qt
from PySide6.QtGui import QPixmap, QDesktopServices, QColor
from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QMessageBox,
    QDialog, QLineEdit, QFormLayout, QInputDialog, QLabel, QHBoxLayout,
    QTableWidget, QTableWidgetItem, QHeaderView, QComboBox, QCompleter,
    QListWidget, QListWidgetItem, QFileDialog, QTabWidget, QDialogButtonBox,
    QTextEdit, QCheckBox, QPlainTextEdit, QProgressDialog
)
from tkinter import Tk, filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import SessionNotCreatedException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

CONFIG_FILE = "config.json"

def criar_csv_vazio(caminho, cabecalhos):
    with open(caminho, mode='w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(cabecalhos)

def verificar_ou_criar_planilhas(diretorio):
    caminho_clientes = os.path.join(diretorio, "clientes.csv")
    caminho_participantes = os.path.join(diretorio, "participantes.csv")

    if not os.path.exists(caminho_clientes):
        cabecalho_clientes = [
            "codigo", "nome_cliente", "email_contador", "email_secundario", "status",
            "pix_pdv", "pix_off", "pos_adiquirente", "boleto", "tef", "delivery", "prioridade"
        ]
        criar_csv_vazio(caminho_clientes, cabecalho_clientes)
        print(f"Arquivo 'clientes.csv' criado em: {caminho_clientes}")

    if not os.path.exists(caminho_participantes):
        cabecalho_participantes = [
            "codigo", "nome", "cod_pais", "cnpj", "cod_mun", "logradouro", "SN",
            "bairro", "endereco", "nome_mun", "prioridade"
        ]
        criar_csv_vazio(caminho_participantes, cabecalho_participantes)
        print(f"Arquivo 'participantes.csv' criado em: {caminho_participantes}")

def selecionar_ou_carregar_diretorio():
    config = {}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
        except UnicodeDecodeError:
            try:
                with open(CONFIG_FILE, "r", encoding="latin1") as f:
                    config = json.load(f)
            except Exception as e:
                print(f"Erro ao ler config.json: {e}")

    caminho = config.get("diretorio_sped")
    if caminho and os.path.isdir(caminho):
        verificar_ou_criar_planilhas(caminho)
        return caminho

    root = Tk()
    root.withdraw()
    pasta = filedialog.askdirectory(title="Selecione a pasta 'CSV SPED'")
    if pasta:
        config["diretorio_sped"] = pasta
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"Erro ao salvar config.json: {e}")
        verificar_ou_criar_planilhas(pasta)
        return pasta
    return None

diretorio_csv = selecionar_ou_carregar_diretorio()
print("Diretório selecionado:", diretorio_csv)

def obter_diretorio_sped():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
        except UnicodeDecodeError:
            try:
                with open(CONFIG_FILE, "r", encoding="latin1") as f:
                    config = json.load(f)
            except Exception as e:
                print(f"Erro ao ler config.json com latin1: {e}")
                config = None
        except Exception as e:
            print(f"Erro ao ler config.json com utf-8: {e}")
            config = None

        if config:
            caminho = config.get("diretorio_sped")
            if caminho and os.path.isdir(caminho):
                return caminho

    caminho = selecionar_ou_carregar_diretorio()

    if caminho:
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump({"diretorio_sped": caminho}, f, indent=4)
        except Exception as e:
            print(f"Erro ao salvar config.json: {e}")
        return caminho
    return None

DIRETORIO_SPED = obter_diretorio_sped()

# Arquivos principais
ARQUIVO_PARTICIPANTES = os.path.join(DIRETORIO_SPED, "participantes.csv")
ARQUIVO_CLIENTES      = os.path.join(DIRETORIO_SPED, "clientes.csv")

# Imagem de cabeçalho
if getattr(sys, 'frozen', False):
    SCRIPT_DIR = sys._MEIPASS
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

HEADER_IMAGE = os.path.join(SCRIPT_DIR, "softcom.jpg")

# Config Partner JSON leitura/escrita com fallback

def carregar_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except UnicodeDecodeError:
            with open(CONFIG_FILE, "r", encoding="latin1") as f:
                return json.load(f)
        except Exception as e:
            print(f"Erro ao ler config.json: {e}")
    return {}

def salvar_config(config):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Erro ao salvar config.json: {e}")

# Extrai base HTTP://IP:PORTA de link
def carregar_partner_config():
    config = carregar_config()
    return config.get("partner_urls", {"urls": []})

def salvar_partner_config(data):
    config = carregar_config()
    config["partner_urls"] = data
    salvar_config(config)

def extrair_base_url(link):
    m = re.match(r'(https?://[\d\.]+:\d+)', link.strip())
    return m.group(1) if m else None

def adicionar_ou_atualizar_usuario(nome, email, senha):
    config = carregar_config()
    if "usuarios" not in config:
        config["usuarios"] = {}
    config["usuarios"][nome.lower()] = {
        "email": email,
        "senha": senha
    }
    salvar_config(config)

def obter_usuario(nome):
    config = carregar_config()
    usuarios = config.get("usuarios", {})
    return usuarios.get(nome.lower())

# Função global reutilizável (fora da classe)
def obter_url_valida(cliente_id):
    config = carregar_config()  # use a que faz sentido para você
    urls = config.get("urls", [])

    for base in urls:
        url_teste = f"{base}/area-partner/public/cliente/index/detail/id/{cliente_id}"
        try:
            response = requests.head(url_teste, timeout=3)
            if response.status_code == 200:
                return url_teste
        except:
            continue

    # Nenhuma funcionou, pede novo link
    novo_link, ok = QInputDialog.getText(
        None,
        "Nova URL do Partner",
        "Nenhuma URL funcionou.\nCole o link completo que você usaria no navegador:"
    )
    if ok and novo_link:
        base_extraida = extrair_base_url(novo_link)
        if base_extraida:
            url_teste = f"{base_extraida}/area-partner/public/cliente/index/detail/id/{cliente_id}"
            try:
                response = requests.head(url_teste, timeout=3)
                if response.status_code == 200:
                    if base_extraida not in urls:
                        urls.insert(0, base_extraida)
                        config["urls"] = urls
                        salvar_config(config)  # ajuste aqui para usar a função correta de salvar
                    return url_teste
            except:
                pass
    return None

class SPED1601GUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('SPED Fiscal')
        self.setFixedSize(400,300)
        self.diretorio_csv = DIRETORIO_SPED
        self.arquivo_participantes = ARQUIVO_PARTICIPANTES
        self.participantes = self.carregar_participantes()
        self.clientes = self.carregar_clientes(ARQUIVO_CLIENTES)

        layout = QVBoxLayout(self)
        if os.path.exists(HEADER_IMAGE):
            lbl = QLabel()
            pix = QPixmap(HEADER_IMAGE).scaled(self.width(),150,Qt.KeepAspectRatio,Qt.SmoothTransformation)
            lbl.setPixmap(pix)
            lbl.setAlignment(Qt.AlignCenter)
            layout.addWidget(lbl)

        botoes = [
            ('Controle SPED', self.controle_SPED),
            ('Inserir Registro 1601', self.inserir_registro_1601),
            ('Limpar Registros 1601', self.limpar_registros_1601),
            ('Enviar E-mail ao Contador', self.enviar_email_contador),
            ('Sair', self.close)
        ]
        for txt, fn in botoes:
            btn = QPushButton(txt)
            btn.clicked.connect(fn)
            layout.addWidget(btn)

    def _salvar_cliente(self, dlg, cliente, inputs):
        for k, w in inputs.items():
            cliente[k] = w.text().strip()
        self.salvar_clientes()
        # arquivo local fica sincronizado pelo Google Drive para desktop
        dlg.accept()

    # =====================================================================================
    # SUAS IMPLEMENTAÇÕES ORIGINAIS DE carregamento de participantes e clientes
    # =====================================================================================

    def _registrar_pendencia_alterar_status(self, cliente_codigo, novo_status):
        """Registra uma pendência apenas de alteração de status no CSV de pendências."""

        # Verifica se o cliente existe na lista atual
        cliente = next((c for c in self.clientes if c.get('codigo') == cliente_codigo), None)
        if not cliente:
            QMessageBox.warning(self, "Erro", "Cliente não encontrado para salvar pendência de status.")
            return

        # Clona os dados do cliente e altera o status
        cliente_pendente = cliente.copy()
        cliente_pendente['status'] = novo_status

        # Salva nos mesmos campos que outras pendências
        self._salvar_pendencia(cliente_pendente)

    def _lock_path(self):
        return os.path.join(self.diretorio_csv, "usuario_atual.json")

    def _pendencias_path(self):
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        return os.path.join(desktop, "pendencias_usuario.csv")
    
    def _adquirir_lock(self):
        lock_path = self._lock_path()
        if os.path.exists(lock_path):
            return False  # lock já existe
        
        try:
            usuario = getpass.getuser()
            maquina = socket.gethostname()
            data_hora = datetime.now().strftime("%H:%M - %d/%m/%Y")
            
            dados_lock = {
                "usuario": usuario,
                "maquina": maquina,
                "data": data_hora
            }
            
            with open(lock_path, 'w', encoding='utf-8') as f:
                json.dump(dados_lock, f, indent=4, ensure_ascii=False)
            
            return True
        except Exception as e:
            print("Erro ao adquirir lock:", e)
            return False

    def _liberar_lock(self):
        lock_path = self._lock_path()
        if os.path.exists(lock_path):
            os.remove(lock_path)

    def _meu_lock(self):
        lock_path = self._lock_path()
        if not os.path.exists(lock_path):
            return False
        
        try:
            with open(lock_path, 'r', encoding='utf-8') as f:
                dados_lock = json.load(f)
            
            usuario = getpass.getuser()
            maquina = socket.gethostname()
            
            return (dados_lock.get("usuario") == usuario and dados_lock.get("maquina") == maquina)
        except Exception:
            return False

    def _ler_lock_info(self):
        lock_path = self._lock_path()
        if not os.path.exists(lock_path):
            return None
        try:
            with open(lock_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return None

    def mostrar_dono_do_lock(self):
        lock_info = self._ler_lock_info()
        if not lock_info:
            print("Nenhum lock encontrado para mostrar.")
            return

        usuario_lock = lock_info.get("usuario", "Desconhecido")
        maquina_lock = lock_info.get("maquina", "Desconhecida")
        data_lock = lock_info.get("data", "Data desconhecida")

        print(f"Mostrando lock: usuário={usuario_lock}, máquina={maquina_lock}, data={data_lock}")

        QMessageBox.information(
            self,
            "Planilha em uso",
            f"🚫 A planilha está atualmente em uso por:\n\n"
            f"👤 Usuário: {usuario_lock}\n"
            f"💻 Máquina: {maquina_lock}\n"
            f"🕒 Desde: {data_lock}"
        )
        
    def _salvar_pendencia(self, cliente):
        """Salva alterações do cliente em pendências."""
        pend_path = os.path.join(os.path.expanduser("~"), "Desktop", "pendencias_usuario.csv")
        os.makedirs(os.path.dirname(pend_path), exist_ok=True)

        campos = [
            'codigo', 'nome_cliente', 'email_contador', 'email_secundario', 'status',
            'pix_pdv', 'pix_off', 'pos_adiquirente', 'boleto', 'tef', 'delivery', 'prioridade'
        ]

        # LIMPEZA: garante que só vai gravar os campos válidos
        cliente_limpo = {campo: cliente.get(campo, '') for campo in campos}

        novo_arquivo = not os.path.exists(pend_path)
        with open(pend_path, mode='a', newline='', encoding='latin1') as f:
            writer = csv.DictWriter(f, fieldnames=campos)
            if novo_arquivo:
                writer.writeheader()
            writer.writerow(cliente_limpo)

        self._remover_arquivo_se_vazio(pend_path)

    def _contar_pendencias(self):
        pend_path = self._pendencias_path()
        if not os.path.exists(pend_path):
            return 0
        with open(pend_path, encoding="latin1") as f:
            return sum(1 for _ in csv.DictReader(f))
    
    def atualizar_contador_pendencias(self):
        total = self._contar_pendencias()
        self.btn_pendencias.setText(f"⚠️ Visualizar pendências: {total}")

    def _remover_arquivo_se_vazio(self, caminho_arquivo):
        if not os.path.exists(caminho_arquivo):
            return
        with open(caminho_arquivo, newline='', encoding='latin1') as f:
            reader = csv.reader(f)
            linhas = list(reader)
            # Se só tiver cabeçalho ou estiver vazio, remove o arquivo
            if len(linhas) <= 1:
                try:
                    os.remove(caminho_arquivo)
                    print(f"Arquivo {caminho_arquivo} removido por estar vazio.")
                except Exception as e:
                    print(f"Erro ao remover arquivo {caminho_arquivo}: {e}")

    def _abrir_pendencias(self):
        pend_path = self._pendencias_path()
        pendencias = []
        if os.path.exists(pend_path):
            with open(pend_path, encoding="latin1") as f:
                pendencias = list(csv.DictReader(f))

        dialog = QDialog(self)
        dialog.setWindowTitle("Pendências do Usuário")
        dialog.setMinimumSize(400, 300)
        layout = QVBoxLayout(dialog)

        lista = QListWidget()
        for p in pendencias:
            lista.addItem(f"{p.get('codigo','')} - {p.get('nome_cliente','')}")
        layout.addWidget(lista)

        # Botão de remover pendência
        btn_remover = QPushButton("Remover Pendência Selecionada")
        btn_remover.setEnabled(False)
        layout.addWidget(btn_remover)

        def on_item_selection():
            btn_remover.setEnabled(len(lista.selectedItems()) > 0)

        lista.itemSelectionChanged.connect(on_item_selection)

        # Botão de importar pendências
        btn_importar = QPushButton("⬇️ Importar Pendências")
        btn_importar.setEnabled(self.edicao_liberada)
        layout.addWidget(btn_importar)

        def importar_e_atualizar():
            self._importar_pendencias(dialog, pend_path)
            self.clientes = self.carregar_clientes(self.caminho_clientes)
            self.participantes = self.carregar_participantes()

        btn_importar.clicked.connect(importar_e_atualizar)

        # Função pra remover pendência selecionada
        def remover_pendencia():
            item = lista.currentItem()
            if not item:
                return

            resposta = QMessageBox.question(
                dialog,
                "Confirmação",
                "Deseja remover a pendência selecionada?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if resposta == QMessageBox.StandardButton.Yes:
                texto = item.text()
                codigo = texto.split(" - ")[0]
                pendencia = next((p for p in pendencias if p['codigo'] == codigo), None)
                if pendencia:
                    pendencias.remove(pendencia)
                    # Regrava o CSV com as pendências atualizadas
                    try:
                        # Aqui fecha o arquivo depois de escrever (with já fecha automaticamente)
                        with open(pend_path, 'w', newline='', encoding='latin1') as f:
                            writer = csv.DictWriter(f, fieldnames=pendencia.keys())
                            writer.writeheader()
                            writer.writerows(pendencias)
                        lista.takeItem(lista.row(item))
                        QMessageBox.information(dialog, "Sucesso", "Pendência removida com sucesso.")

                        # Agora tente remover o arquivo se estiver vazio, só após o 'with' fechar o arquivo
                        self._remover_arquivo_se_vazio(pend_path)
                        self.atualizar_contador_pendencias()
                        
                    except Exception as e:
                        QMessageBox.warning(dialog, "Erro", f"Erro ao remover pendência: {e}")

        btn_remover.clicked.connect(remover_pendencia)
        self.atualizar_contador_pendencias()

        dialog.exec()

    def _importar_pendencias(self, dialog, pend_path):
        if not self.edicao_liberada:
            QMessageBox.warning(self, "Bloqueado", "A planilha está em uso por outro usuário.")
            return
        if os.path.exists(pend_path):
            with open(pend_path, encoding="latin1") as f:
                pendencias = list(csv.DictReader(f))
            # Cria um dicionário para acesso rápido pelo código
            clientes_dict = {c['codigo']: c for c in self.clientes}
            for pend in pendencias:
                codigo = pend['codigo']
                if codigo in clientes_dict:
                    # Atualiza todos os campos do cliente existente
                    clientes_dict[codigo].update(pend)
                else:
                    # Adiciona novo cliente se não existir
                    self.clientes.append(pend)
            self.salvar_clientes()
            os.remove(pend_path)
            self._remover_arquivo_se_vazio(self._pendencias_path())
            QMessageBox.information(self, "Importado", "Pendências importadas e mescladas com sucesso!")
            if dialog is not None:
                dialog.accept()

    def abrir_participante(self, row, column):
        participante = self.participantes[row]
        dialog = FormDetalhesParticipante(participante, self)
        dialog.exec()

    def carregar_participantes(self):
        caminho_participantes = os.path.join(DIRETORIO_SPED, "participantes.csv")
        if not os.path.exists(caminho_participantes):
            print("Arquivo 'participantes.csv' não encontrado!")
            return []

        try:
            with open(caminho_participantes, mode='r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                participantes = [row for row in reader]
        except UnicodeDecodeError:
            # Tenta latin1 se utf-8 falhar
            with open(caminho_participantes, mode='r', encoding='latin1') as f:
                reader = csv.DictReader(f)
                participantes = [row for row in reader]
        return participantes

    def carregar_clientes(self, caminho):
        """Carrega os clientes de um arquivo CSV no diretório SPED."""
        try:
            if not os.path.exists(caminho):
                return
            try:
                with open(caminho, mode='r', encoding='utf-8') as arquivo:
                    reader = csv.DictReader(arquivo)
                    self.clientes = [linha for linha in reader]
                    for cliente in self.clientes:
                        if 'prioridade' not in cliente:
                            cliente['prioridade'] = 'False'                    

            except UnicodeDecodeError:
                with open(caminho, mode='r', encoding='latin1') as arquivo:
                    reader = csv.DictReader(arquivo)
                    self.clientes = [dict(linha) for linha in reader if linha and any(linha.values())]
                    for cliente in self.clientes:
                        if 'prioridade' not in cliente:
                            cliente['prioridade'] = 'False'
        except Exception as e:
            QMessageBox.warning(self, "Erro", f"Erro ao carregar clientes: {e}")

    def listar_clientes(self):
        clientes_path = os.path.join(DIRETORIO_SPED, "clientes.csv")
        participantes_path = os.path.join(DIRETORIO_SPED, "participantes.csv")

        # Verifica existência dos arquivos
        if not os.path.exists(clientes_path):
            QMessageBox.information(self, "Arquivo não encontrado", "Arquivo 'clientes.csv' não encontrado.")
            return
        if not os.path.exists(participantes_path):
            QMessageBox.information(self, "Arquivo não encontrado", "Arquivo 'participantes.csv' não encontrado.")
            return

        # Carrega dados de clientes e participantes
        with open(clientes_path, mode='r', encoding='latin1') as f:
            self.clientes = list(csv.DictReader(f))
        with open(participantes_path, mode='r', encoding='latin1') as f:
            participantes = list(csv.DictReader(f))

        # Monta diálogo para seleção de cliente
        dialog_cliente = FormSelecionarCliente(self.clientes, self)

        # Conecta duplo clique para abrir dados do cliente
        dialog_cliente.list_widget.itemDoubleClicked.connect(lambda item: self.abrir_dados_cliente(item, dialog_cliente, participantes, clientes_path))

        if dialog_cliente.exec() == QDialog.Accepted:
            selected_item = dialog_cliente.list_widget.currentItem()
            if not selected_item:
                QMessageBox.information(self, "Informação", "Nenhum cliente selecionado.")
                return
            cliente = selected_item.data(Qt.UserRole)

            # Monta texto de detalhes do cliente selecionado (exemplo original)
            texto = f"=== Cliente Selecionado: {cliente['nome_cliente']} ===\n\n"
            texto += f"Código: {cliente['codigo']}\n"
            texto += f"E-mail do contador: {cliente['email_contador']}\n"
            texto += f"E-mail secundário: {cliente.get('email_secundario', '')}\n"
            texto += f"Status: {cliente['status']}\n"
            texto += "Meios de Pagamento:\n"
            for metodo in ["pix_pdv", "pix_off", "pos_adiquirente", "boleto", "tef", "delivery"]:
                banco_codigo = cliente.get(metodo, "")
                if banco_codigo:
                    banco_info = next((p for p in participantes if p['cnpj'] == banco_codigo), None)
                    if banco_info:
                        nome_banco = banco_info.get('nome', 'Não Informado')
                        nome_mun = banco_info.get('nome_mun', 'Não Informado')
                        texto += f"  - {metodo.upper()}: Banco {nome_banco} (CNPJ: {banco_codigo}) - Município: {nome_mun}\n"

            self.show_text_dialog("Detalhes do Cliente", texto)
        else:
            return

        # Oferece opção de ver clientes com status vazio
        resposta = QMessageBox.question(
            self,
            "Clientes com status vazio",
            "Deseja ver os clientes que faltam gerar o SPED (status vazio)?",
            QMessageBox.Yes | QMessageBox.No
        )
        if resposta == QMessageBox.Yes:
            faltando = [c for c in self.clientes if c['status'] == ""]
            if not faltando:
                QMessageBox.information(self, "Clientes com status vazio", "Não há clientes com status vazio.")
            else:
                texto_faltando = "=== Clientes com Status Vazio ===\n"
                for i, c in enumerate(faltando, 1):
                    texto_faltando += f"\n{i}. Código: {c['codigo']} - Nome: {c['nome_cliente']} - E-mail: {c['email_contador']}\n"
                self.show_text_dialog("Clientes com Status Vazio", texto_faltando)

    def abrir_dados_cliente(self, item, dialog_cliente, participantes):
        cliente = item.data(Qt.UserRole)

        dialog_cliente_editar = QDialog(dialog_cliente)
        dialog_cliente_editar.setWindowTitle(f"Dados do Cliente: {cliente.get('nome_cliente', '')}")
        dialog_cliente_editar.setModal(True)
        layout = QVBoxLayout()
        dialog_cliente_editar.setLayout(layout)

        campos = [
            ('Código', 'codigo'),
            ('Nome do Cliente', 'nome_cliente'),
            ('E-mail do Contador', 'email_contador'),
            ('E-mail Secundário', 'email_secundario'),
            ('Status', 'status'),
        ]

        widgets = {}

        for label_text, chave in campos:
            hbox = QHBoxLayout()
            lbl = QLabel(label_text + ":")
            edt = QLineEdit(cliente.get(chave, ''))
            hbox.addWidget(lbl)
            hbox.addWidget(edt)
            layout.addLayout(hbox)
            widgets[chave] = edt

        btn_salvar = QPushButton("Salvar")
        layout.addWidget(btn_salvar)

        def salvar_alteracoes():
            for chave, widget in widgets.items():
                cliente[chave] = widget.text()

            try:
                with open(ARQUIVO_CLIENTES, mode='w', encoding='latin1', newline='') as f:
                    writer = csv.DictWriter(f, fieldnames=self.clientes[0].keys())
                    writer.writeheader()
                    writer.writerows(self.clientes)
                QMessageBox.information(dialog_cliente_editar, "Sucesso", "Dados do cliente atualizados com sucesso!")
                dialog_cliente_editar.accept()
            except Exception as e:
                QMessageBox.critical(dialog_cliente_editar, "Erro", f"Falha ao salvar os dados: {e}")

        btn_salvar.clicked.connect(salvar_alteracoes)
        dialog_cliente_editar.exec_()

    def abrir_dados_participante(self, item, dialog_participantes, participantes_path):
        row = item.row() if hasattr(item, 'row') else item  # se item for índice da linha
        participante = self.participantes[row]

        dialog_participante_editar = QDialog(dialog_participantes)
        dialog_participante_editar.setWindowTitle(f"Dados do Participante: {participante.get('nome', '')}")
        dialog_participante_editar.setModal(True)
        layout = QVBoxLayout()
        dialog_participante_editar.setLayout(layout)

        campos = [
            ('Nome', 'nome'),
            ('CNPJ', 'cnpj'),
            ('Código do País', 'cod_pais'),
            ('Código do Município', 'cod_mun'),
            ('Nome do Município', 'nome_mun'),
            ('Logradouro', 'logradouro'),
            ('Bairro', 'bairro'),
            ('Número', 'SN'),
        ]

        widgets = {}

        for label_text, chave in campos:
            hbox = QHBoxLayout()
            lbl = QLabel(label_text + ":")
            edt = QLineEdit(participante.get(chave, ''))
            hbox.addWidget(lbl)
            hbox.addWidget(edt)
            layout.addLayout(hbox)
            widgets[chave] = edt

        btn_salvar = QPushButton("Salvar")
        layout.addWidget(btn_salvar)

        def salvar_alteracoes():
            for chave, widget in widgets.items():
                participante[chave] = widget.text()

            try:
                with open(participantes_path, mode='w', encoding='latin1', newline='') as f:
                    writer = csv.DictWriter(f, fieldnames=self.participantes[0].keys())
                    writer.writeheader()
                    writer.writerows(self.participantes)
                QMessageBox.information(dialog_participante_editar, "Sucesso", "Dados do participante atualizados com sucesso!")
                dialog_participante_editar.accept()
            except Exception as e:
                QMessageBox.critical(dialog_participante_editar, "Erro", f"Falha ao salvar os dados: {e}")

        btn_salvar.clicked.connect(salvar_alteracoes)

        dialog_participante_editar.exec_()

    def filtrar_tabela_por_status(self, status_permitidos, filtrar_prioridade=False):
        col_prioridade = self.tabela_clientes.columnCount() - 1
        status_permitidos_upper = {s.upper() for s in status_permitidos}

        for row in range(self.tabela_clientes.rowCount()):
            item_status = self.tabela_clientes.item(row, 4)
            status = item_status.text().strip().upper() if item_status else ''

            item_prio = self.tabela_clientes.item(row, col_prioridade)
            prioridade_eh_vermelha = False
            if item_prio:
                cor = item_prio.background().color()
                prioridade_eh_vermelha = (
                    cor == QColor('red') or
                    cor.name().lower() == '#ff0000'
                )

            # Lógica de exibição
            if not status_permitidos and not filtrar_prioridade:
                mostrar = True  # mostrar tudo
            else:
                mostrar = False
                if status_permitidos_upper:
                    mostrar = status in status_permitidos_upper
                if filtrar_prioridade:
                    mostrar = mostrar or prioridade_eh_vermelha

            self.tabela_clientes.setRowHidden(row, not mostrar)

        # Atualizar contador de visíveis
        visiveis = sum(
            not self.tabela_clientes.isRowHidden(row)
            for row in range(self.tabela_clientes.rowCount())
        )
        self.lbl_visiveis.setText(f"<b>Clientes:</b> {visiveis}")

    # =====================================================================================
    # NOVO: método controle_SPED() – cria o painel administrativo em abas
    # =====================================================================================

    def mostrar_todos_clientes(self):
        # Mostrar todas as linhas, sem filtro (desocultar tudo)
        for row in range(self.tabela_clientes.rowCount()):
            self.tabela_clientes.setRowHidden(row, False)
        # Atualizar contador de visíveis
        visiveis = self.tabela_clientes.rowCount()
        self.lbl_visiveis.setText(f"<b>Clientes:</b> {visiveis}")

    def abrir_filtro_status(self):
        status_unicos = sorted(set(
            (c.get('status') or '').upper()
            for c in self.clientes if c.get('status')
        ))

        dialog = QDialog(self)
        dialog.setWindowTitle("Filtrar por Status")
        layout = QVBoxLayout(dialog)

        # Checkbox Prioridade (independente dos status)
        cb_prioridade = QCheckBox("PRIORIDADE")
        cb_prioridade.setChecked(False)
        layout.addWidget(cb_prioridade)

        # Checkboxes para status
        checkboxes = {}
        for status in status_unicos:
            cb = QCheckBox(status)
            cb.setChecked(False)
            layout.addWidget(cb)
            checkboxes[status] = cb

        botoes_layout = QHBoxLayout()
        btn_aplicar = QPushButton("🎯 Aplicar filtro")
        btn_limpar = QPushButton("❌ Limpar filtro")
        botoes_layout.addWidget(btn_aplicar)
        botoes_layout.addWidget(btn_limpar)
        layout.addLayout(botoes_layout)

        def aplicar():
            prioridade_marcada = cb_prioridade.isChecked()
            status_marcados = {k for k, cb in checkboxes.items() if cb.isChecked()}

            if not status_marcados and not prioridade_marcada:
                # Nenhum filtro, mostrar todos
                self.mostrar_todos_clientes()
            else:
                # Chama função de filtro com os parâmetros
                self.filtrar_tabela_por_status(status_marcados, filtrar_prioridade=prioridade_marcada)

            dialog.accept()

        def limpar():
            cb_prioridade.setChecked(False)
            for cb in checkboxes.values():
                cb.setChecked(False)
            self.mostrar_todos_clientes()
            dialog.accept()

        btn_aplicar.clicked.connect(aplicar)
        btn_limpar.clicked.connect(limpar)

        dialog.exec_()

    def controle_SPED(self):
        # Tenta adquirir o lock
        self.edicao_liberada = self._adquirir_lock()

        # Só agora pode carregar os dados
        self.caminho_clientes = os.path.join(DIRETORIO_SPED, "clientes.csv")
        self.carregar_clientes(self.caminho_clientes)
        self.participantes = self.carregar_participantes()

        self.hide()
        dialog = QDialog()
        dialog.setWindowTitle("Controle SPED - Painel Administrativo")
        dialog.setWindowState(dialog.windowState() | Qt.WindowMaximized)
        dialog.setWindowFlags(dialog.windowFlags() | Qt.WindowMinMaxButtonsHint)

        layout_principal = QVBoxLayout()
        dialog.setLayout(layout_principal)

        # Cria o QTabWidget
        tabs = QTabWidget()
        layout_principal.addWidget(tabs)

        # === Aba “Clientes” ===
        aba_clientes = QWidget()
        layout_clientes = QVBoxLayout()
        aba_clientes.setLayout(layout_clientes)
    
        # 1) Container superior para botões (linha única)
        container_botoes = QWidget()
        botoes_layout = QHBoxLayout()
        botoes_layout.setContentsMargins(0, 0, 0, 0)
        botoes_layout.setSpacing(10)  # pequeno espaçamento entre botões
        container_botoes.setLayout(botoes_layout)

        # Botão “Cadastrar Cliente”
        btn_cadastrar = QPushButton("➕ Cadastrar Cliente")
        botoes_layout.addWidget(btn_cadastrar, alignment=Qt.AlignLeft)
        btn_cadastrar.setStyleSheet("font-size: 14px;")

        # Botão “Remover Cliente”
        btn_remover = QPushButton("🗑️ Remover Cliente")
        botoes_layout.addWidget(btn_remover, alignment=Qt.AlignLeft)
        btn_remover.setStyleSheet("font-size: 14px;")

        # Botão “Alterar Dados”
        btn_alterar = QPushButton("✏️ Alterar Dados")
        botoes_layout.addWidget(btn_alterar, alignment=Qt.AlignLeft)
        btn_alterar.setStyleSheet("font-size: 14px;")

        # Botão "Zerar status"
        btn_zerar_status = QPushButton("🔄 Zerar status")
        botoes_layout.addWidget(btn_zerar_status, alignment=Qt.AlignLeft)
        btn_zerar_status.setStyleSheet("font-size: 14px;")

        # Botão "Filtrar status"
        btn_filtrar_status = QPushButton("🔍 Filtrar status")
        btn_filtrar_status.clicked.connect(self.abrir_filtro_status)
        botoes_layout.addWidget(btn_filtrar_status, alignment=Qt.AlignLeft)
        btn_filtrar_status.setStyleSheet("font-size: 14px;")

        # Botão "Visualizar pendências"
        self.btn_pendencias = QPushButton(f"Visualizar pendências ({self._contar_pendencias()})")
        self.btn_pendencias.clicked.connect(self._abrir_pendencias)
        botoes_layout.addWidget(self.btn_pendencias, alignment=Qt.AlignLeft)
        self.btn_pendencias.setStyleSheet("font-size: 14px;")
        self.atualizar_contador_pendencias()

        # Adiciona espaço para empurrar os botões à esquerda
        botoes_layout.addStretch()

        # Adiciona o container ao layout principal
        layout_clientes.addWidget(container_botoes)

        # 2) Campo de filtro (abaixo da linha de botões)
        filtro_clientes = QLineEdit()
        filtro_clientes.setPlaceholderText("Filtrar por nome ou código")
        filtro_clientes.setFixedWidth(250)
        layout_clientes.addWidget(filtro_clientes)

        # 3) Estatísticas de clientes (abaixo do filtro)
        total_clientes = len(self.clientes)
        feitos = sum(1 for c in self.clientes if (c.get('status') or '').upper() == 'FEITO')
        pendentes = total_clientes - feitos

        # Criando labels individuais
        self.lbl_total = QLabel(f"<b>Total de clientes:</b> {total_clientes}")
        self.lbl_total.setWordWrap(False)

        self.lbl_feitos = QLabel(f"<b>SPEDs já gerados (FEITO):</b> {feitos}")
        self.lbl_feitos.setStyleSheet("color: green;")
        self.lbl_feitos.setWordWrap(False)

        self.lbl_pendentes = QLabel(f"<b>Faltam SPEDs (PENDENTE):</b> {pendentes}")
        self.lbl_pendentes.setStyleSheet("color: red;")
        self.lbl_pendentes.setWordWrap(False)

        self.lbl_visiveis = QLabel()  # será atualizado após preencher a tabela
        self.lbl_visiveis.setStyleSheet("color: orange;")
        self.lbl_visiveis.setWordWrap(False)

        # Criar layout horizontal para pendentes + visíveis na mesma linha
        hlayout_pendentes = QHBoxLayout()
        hlayout_pendentes.addWidget(self.lbl_pendentes)
        hlayout_pendentes.addStretch()
        hlayout_pendentes.addWidget(self.lbl_visiveis)

        # Adicionando as labels no layout (um abaixo do outro)
        layout_clientes.addWidget(self.lbl_total)
        layout_clientes.addWidget(self.lbl_feitos)
        layout_clientes.addLayout(hlayout_pendentes)

        # 4) Tabela de clientes
        colunas = [
            "Código",
            "Nome do cliente",
            "E-mail do contador",
            "E-mail secundário",
            "Status",
            "PIX PDV",
            "PIX Off",
            "POS adquirente",
            "Boleto",
            "TEF",
            "Delivery"
        ]
        self.tabela_clientes = QTableWidget()
        self.tabela_clientes.setColumnCount(len(colunas))
        self.tabela_clientes.setHorizontalHeaderLabels(colunas)
        self.tabela_clientes.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tabela_clientes.setSelectionBehavior(QTableWidget.SelectRows)
        self.tabela_clientes.setAlternatingRowColors(True)
        layout_clientes.addWidget(self.tabela_clientes, stretch=1)

        # Função auxiliar para converter CNPJ em nome de banco
        def nome_do_banco_por_cnpj(cnpj):
            cnpj = (cnpj or "").strip()
            if not cnpj:
                return ""
            participante = next(
                (p for p in self.participantes if p.get('cnpj', '').strip() == cnpj),
                None
            )
            return participante.get('nome', '') if participante else ""

        # Função interna que atualiza estatísticas e preenche a tabela
        def atualizar_tabela_clientes():
            # Recarregar clientes do CSV
            path_clientes = os.path.join(self.diretorio_csv, "clientes.csv")
            self.carregar_clientes(path_clientes)

            # Atualizar estatísticas
            total = len(self.clientes)
            feitos_local = sum(1 for c in self.clientes if (c.get('status') or '').upper() == 'FEITO')
            pendentes_local = sum(1 for c in self.clientes if (c.get('status') or '').upper() == 'PENDENTE')

            # Atualizar labels
            self.lbl_total.setText(f"<b>Total de clientes:</b> {total}")
            self.lbl_feitos.setText(f"<b>SPEDs já gerados (FEITO):</b> {feitos_local}")
            self.lbl_pendentes.setText(f"<b>Faltam SPEDs (PENDENTE):</b> {pendentes_local}")

            # Atualizar tabela
            self.tabela_clientes.setRowCount(total)

            # Atualizar quantidade de clientes visíveis (não escondidos)
            visiveis = sum(
                not self.tabela_clientes.isRowHidden(row)
                for row in range(self.tabela_clientes.rowCount())
            )
            self.lbl_visiveis.setText(f"<b>Clientes:</b> {visiveis}")

            # Preencher cada célula
            for row_idx, cli in enumerate(self.clientes):
                # Cria os QTableWidgetItem com flags para não permitir edição
                def criar_item(texto):
                    item = QTableWidgetItem(texto)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)  # remove a flag de edição
                    return item

                self.tabela_clientes.setItem(row_idx, 0, criar_item(cli.get('codigo', '')))
                self.tabela_clientes.setItem(row_idx, 1, criar_item(cli.get('nome_cliente', '')))
                self.tabela_clientes.setItem(row_idx, 2, criar_item(cli.get('email_contador', '')))
                self.tabela_clientes.setItem(row_idx, 3, criar_item(cli.get('email_secundario', '')))
                self.tabela_clientes.setItem(row_idx, 4, criar_item(cli.get('status', '')))

                self.tabela_clientes.setItem(row_idx, 5, criar_item(nome_do_banco_por_cnpj(cli.get('pix_pdv', ''))))
                self.tabela_clientes.setItem(row_idx, 6, criar_item(nome_do_banco_por_cnpj(cli.get('pix_off', ''))))
                self.tabela_clientes.setItem(row_idx, 7, criar_item(nome_do_banco_por_cnpj(cli.get('pos_adiquirente', ''))))
                self.tabela_clientes.setItem(row_idx, 8, criar_item(nome_do_banco_por_cnpj(cli.get('boleto', ''))))
                self.tabela_clientes.setItem(row_idx, 9, criar_item(nome_do_banco_por_cnpj(cli.get('tef', ''))))
                self.tabela_clientes.setItem(row_idx, 10, criar_item(nome_do_banco_por_cnpj(cli.get('delivery', ''))))

                # Pinta a linha de vermelho se for prioridade
                if (cli.get('prioridade') or '').lower() == 'sim':
                    for col in range(self.tabela_clientes.columnCount()):
                        item = self.tabela_clientes.item(row_idx, col)
                        if item:
                            item.setBackground(QColor(255, 0, 0))  # vermelho semitransparente
                            
            # Após preencher, reaplicar filtro ativo (se houver)
            texto_filtro = filtro_clientes.text().strip().lower()
            aplicar_filtro_clientes(texto_filtro)

        def zerar_status_clientes():
            """
            Altera o campo 'status' de todos os clientes que estão como 'FEITO' para 'PENDENTE',
            e salva no CSV — somente se o lock for do usuário atual.
            """

            # Verifica se o lock é do usuário atual
            if not self._meu_lock():
                QMessageBox.warning(
                    self,
                    "Bloqueado",
                    "Você não tem permissão para alterar os dados.\n"
                    "A planilha está sendo editada por outro usuário."
                )
                return

            # Confirmação
            resp = QMessageBox.question(
                self,
                "Zerar STATUS dos Clientes",
                "Deseja realmente alterar o status de TODOS os clientes de 'FEITO' para 'PENDENTE'?",
                QMessageBox.Yes | QMessageBox.No
            )
            if resp != QMessageBox.Yes:
                return

            # Atualiza status dos clientes em memória
            alterou = False
            for cliente in self.clientes:
                status = cliente.get('status', '').strip().upper()
                if status == 'FEITO':
                    cliente['status'] = 'PENDENTE'
                    alterou = True

            # Só salva e mostra mensagem se realmente alterou algo
            if alterou:
                self.salvar_clientes()
                QMessageBox.information(self, "Sucesso", "Status atualizado de 'FEITO' para 'PENDENTE' com sucesso!")
                atualizar_tabela_clientes()
            else:
                QMessageBox.information(self, "Aviso", "Nenhum cliente com status 'FEITO' foi encontrado.")

        # Botão “Zerar status”
        btn_zerar_status.clicked.connect(zerar_status_clientes)
        
        # Função para mostrar/ocultar linhas conforme texto de filtro
        def aplicar_filtro_clientes(texto):
            for row_idx in range(self.tabela_clientes.rowCount()):
                item_codigo = self.tabela_clientes.item(row_idx, 0)
                item_nome = self.tabela_clientes.item(row_idx, 1)
                codigo_texto = item_codigo.text().lower() if item_codigo else ""
                nome_texto = item_nome.text().lower() if item_nome else ""
                if not texto:
                    self.tabela_clientes.setRowHidden(row_idx, False)
                else:
                    # só mostra se o filtro estiver em código ou em nome
                    if texto in codigo_texto or texto in nome_texto:
                        self.tabela_clientes.setRowHidden(row_idx, False)
                    else:
                        self.tabela_clientes.setRowHidden(row_idx, True)

        # Conecta o filtro para reagir a cada mudança de texto
        filtro_clientes.textChanged.connect(lambda t: aplicar_filtro_clientes(t.strip().lower()))

        # Preenche a tabela pela primeira vez
        atualizar_tabela_clientes()

        # Conecta o clique do botão “Cadastrar Cliente”
        def on_cadastrar_click():
            try:
                self.cadastrar_cliente()
            except PermissionError as e:
                QMessageBox.warning(
                    self,
                    "Permissão Negada",
                    "Não foi possível abrir/alterar 'clientes.csv'.\n"
                    "Verifique se o arquivo está aberto em outro programa ou se você tem permissão de gravação.\n"
                    f"Detalhes: {e}"
                )
                return
            # Após cadastro, atualiza a tabela e reaplica filtro
            atualizar_tabela_clientes()

        btn_cadastrar.clicked.connect(on_cadastrar_click)

        # Conecta o clique do botão “Remover Cliente”
        def on_remover_click():
            try:
                self.remover_cliente()
            except Exception as e:
                QMessageBox.warning(
                    self,
                    "Erro ao Remover",
                    f"Não foi possível remover o cliente.\nDetalhes: {e}"
                )
                return
            # Depois de remover, atualiza a tabela e reaplica filtro
            atualizar_tabela_clientes()

        btn_remover.clicked.connect(on_remover_click)

        # Conecta o clique do botão “Alterar Dados”
        def on_alterar_click():
            try:
                self.trocar_banco_cliente()
            except Exception as e:
                QMessageBox.warning(
                    self,
                    "Erro ao Alterar",
                    f"Não foi possível alterar os dados do cliente.\nDetalhes: {e}"
                )
                return
            # Após alteração, atualiza a tabela e reaplica filtro
            atualizar_tabela_clientes()

        btn_alterar.clicked.connect(on_alterar_click)

        tabs.addTab(aba_clientes, "Clientes")

        # Dentro do método onde você trata o duplo clique no cliente:
        def on_cliente_duplo_clique(item):
            row = item.row()
            cliente = self.clientes[row]

            janela_edicao = QDialog(dialog)
            janela_edicao.setWindowTitle(f"Editar Cliente - {cliente.get('nome_cliente', '')}")
            janela_edicao.setModal(True)

            main_layout = QVBoxLayout(janela_edicao)
            # Cabeçalho com botões
            header_layout = QHBoxLayout()
            btn_partner = QPushButton("Cadastro Partner")
            btn_ocorrencia = QPushButton("Gerar Ocorrência")
            btn_inserir_1601 = QPushButton("Inserir Registro 1601")
            btn_limpar_1601 = QPushButton("Limpar Registro 1601")
            btn_enviar_email = QPushButton("Enviar e-mail ao contador")

            header_layout.addWidget(btn_partner)
            header_layout.addWidget(btn_ocorrencia)
            header_layout.addWidget(btn_inserir_1601)
            header_layout.addWidget(btn_limpar_1601)
            header_layout.addWidget(btn_enviar_email)
            header_layout.addStretch()

            checkbox_prioridade = QCheckBox("Prioridade")
            checkbox_prioridade.setChecked(cliente.get("prioridade", "Não") == "Sim")
            header_layout.addWidget(checkbox_prioridade)

            main_layout.addLayout(header_layout)

            #botão de conectar
            btn_inserir_1601.clicked.connect(lambda: self.inserir_registro_1601_com_cliente(cliente, atualizar_tabela_clientes))
            btn_limpar_1601.clicked.connect(lambda: self.limpar_registros_1601())
            btn_enviar_email.clicked.connect(lambda: self.enviar_email_contador_com_cliente(cliente))

            # Função original de abrir cadastro
            def open_partner():
                cliente_id = cliente.get('codigo', '')
                url_valida = obter_url_valida(cliente_id)
                if url_valida:
                    QDesktopServices.openUrl(QUrl(url_valida))
                else:
                    QMessageBox.warning(
                        janela_edicao,
                        "Erro",
                        "Não foi possível acessar nenhuma URL válida para este cliente."
                    )
            btn_partner.clicked.connect(open_partner)

            def gerar_ocorrencia():
                cliente_id = cliente.get('codigo', '')
                url = obter_url_valida(cliente_id)
                if not url:
                    QMessageBox.warning(
                        None,
                        "Erro",
                        "Não foi possível obter a URL para este cliente."
                    )
                    return

                # --- diálogo interno para escolher usuário + botão Novo ---
                class DialogSelecionarUsuario(QDialog):
                    def __init__(self, parent=None):
                        super().__init__(parent)
                        self.setWindowTitle("Selecionar usuário")
                        self.usuario_selecionado = None

                        layout = QVBoxLayout(self)

                        layout.addWidget(QLabel("Quem está utilizando?"))

                        self.combo_usuarios = QComboBox()
                        self.carregar_usuarios()
                        layout.addWidget(self.combo_usuarios)

                        btn_novo = QPushButton("Novo")
                        btn_novo.clicked.connect(self.novo_usuario)

                        btn_ok = QPushButton("OK")
                        btn_ok.clicked.connect(self.accept)
                        btn_cancel = QPushButton("Cancelar")
                        btn_cancel.clicked.connect(self.reject)

                        hbox = QHBoxLayout()
                        hbox.addWidget(btn_novo)
                        hbox.addStretch()
                        hbox.addWidget(btn_ok)
                        hbox.addWidget(btn_cancel)

                        layout.addLayout(hbox)

                    def carregar_usuarios(self):
                        self.combo_usuarios.clear()
                        config = carregar_config()
                        usuarios = config.get("usuarios", {})
                        nomes = sorted(usuarios.keys())
                        self.combo_usuarios.addItems(nomes)

                    def novo_usuario(self):
                        nome, ok1 = QInputDialog.getText(self, "Novo usuário", "Nome do usuário:")
                        if not ok1 or not nome.strip():
                            return
                        email, ok2 = QInputDialog.getText(self, "Novo usuário", "E-mail:")
                        if not ok2 or not email.strip():
                            return
                        senha, ok3 = QInputDialog.getText(self, "Novo usuário", "Senha:", QLineEdit.Password)
                        if not ok3 or not senha.strip():
                            return

                        adicionar_ou_atualizar_usuario(nome.strip(), email.strip(), senha.strip())
                        QMessageBox.information(self, "Sucesso", f"Usuário '{nome.strip()}' adicionado com sucesso.")
                        self.carregar_usuarios()
                        idx = self.combo_usuarios.findText(nome.strip())
                        if idx >= 0:
                            self.combo_usuarios.setCurrentIndex(idx)

                    def get_usuario(self):
                        return self.combo_usuarios.currentText()

                dialog = DialogSelecionarUsuario(None)
                if dialog.exec_() != QDialog.Accepted:
                    return
                nome_usuario = dialog.get_usuario()
                if not nome_usuario:
                    QMessageBox.warning(janela_edicao, "Erro", "Nenhum usuário selecionado.")
                    return

                try:
                    chrome_options = webdriver.ChromeOptions()
                    chrome_options.add_argument("--start-maximized")
                    driver = webdriver.Chrome(options=chrome_options)

                    wait = WebDriverWait(driver, 20)
                    driver.get(url)

                    fez_login = False

                    config = carregar_config()
                    usuarios = config.get("usuarios", {})
                    usuario_info = usuarios.get(nome_usuario)

                    if not usuario_info:
                        QMessageBox.warning(janela_edicao, "Erro", f"Usuário '{nome_usuario}' não encontrado na configuração.")
                        return

                    email = usuario_info.get("email")
                    senha = usuario_info.get("senha")

                    campo_usuario = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="user"]')))
                    campo_usuario.clear()
                    campo_usuario.send_keys(email)

                    campo_senha = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="senha"]')))
                    campo_senha.clear()
                    campo_senha.send_keys(senha)
                    campo_senha.send_keys(u'\ue007')  # Enter

                    fez_login = True

                except Exception as e:
                    QMessageBox.warning(janela_edicao, "Erro", f"Erro no Selenium: {e}")
                    return

                if fez_login:
                    wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
                    driver.get(url)

                try:
                    xpath_ocorrencia = '/html/body/div[2]/div[2]/ul/li[15]/a'
                    botao = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_ocorrencia)))
                    botao.click()

                    url_ocorrencia = driver.current_url

                    xpath_abrir_ocorrencia = '//*[@id="btn-ocorrencia"]'
                    botao_abrir = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_abrir_ocorrencia)))
                    botao_abrir.click()

                    time.sleep(1)
                    driver.switch_to.window(driver.window_handles[-1])

                    xpath_solicitante = '//*[@id="filter"]/div[2]/div/form/div[1]/input'
                    campo_solicitante = wait.until(EC.presence_of_element_located((By.XPATH, xpath_solicitante)))
                    campo_solicitante.clear()
                    campo_solicitante.send_keys(cliente.get("nome_cliente", ""))

                    assuntos = [
                        "FR TEC SUPORTE REMOTO",
                        "FR TEC SPED",
                        "FR TEC SUPORTE PRESENCIAL"
                    ]
                    assunto, ok = QInputDialog.getItem(
                        None,
                        "Assunto",
                        "Selecione o assunto:",
                        assuntos,
                        editable=False
                    )
                    if not ok:
                        return

                    campo_assunto_container = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="select2-assunto-container"]')))
                    campo_assunto_container.click()
                    time.sleep(1)
                    opcao_assunto = wait.until(EC.element_to_be_clickable((By.XPATH, f'//li[contains(text(), "{assunto}")]')))
                    opcao_assunto.click()

                    xpath_motivo = '//*[@id="filter"]/div[2]/div/form/div[7]/textarea'
                    campo_motivo = wait.until(EC.presence_of_element_located((By.XPATH, xpath_motivo)))
                    if assunto == "FR TEC SPED":
                        campo_motivo.send_keys("Gerar SPED.")
                        motivo_gerado = "Gerar SPED."
                    else:
                        motivo, ok = QInputDialog.getText(None, "Motivo", "Digite o motivo:")
                        if not ok:
                            return
                        campo_motivo.send_keys(motivo)
                        motivo_gerado = motivo

                    xpath_servico = '//*[@id="filter"]/div[2]/div/form/div[16]/textarea'
                    campo_servico = wait.until(EC.presence_of_element_located((By.XPATH, xpath_servico)))
                    if assunto == "FR TEC SPED":
                        campo_servico.send_keys("Gerado e enviado.")
                    else:
                        servico, ok = QInputDialog.getText(None, "Serviço realizado", "Digite o serviço realizado:")
                        if not ok:
                            return
                        campo_servico.send_keys(servico)

                    xpath_chegada = '//*[@id="HoraChegada"]'
                    hora_chegada_raw, ok = QInputDialog.getText(None, "Hora de Chegada", "Digite a hora de chegada (formato HHMM):")
                    if not ok:
                        return
                    hora_chegada = f"{hora_chegada_raw[:2]}:{hora_chegada_raw[2:]}"
                    campo_chegada = wait.until(EC.presence_of_element_located((By.XPATH, xpath_chegada)))
                    campo_chegada.clear()
                    campo_chegada.send_keys(hora_chegada)

                    xpath_saida = '//*[@id="HoraSaida"]'
                    hora_saida_raw, ok = QInputDialog.getText(None, "Hora de Saída", "Digite a hora de saída (formato HHMM):")
                    if not ok:
                        return
                    hora_saida = f"{hora_saida_raw[:2]}:{hora_saida_raw[2:]}"
                    campo_saida = wait.until(EC.presence_of_element_located((By.XPATH, xpath_saida)))
                    campo_saida.clear()
                    campo_saida.send_keys(hora_saida)

                    botao_enviar = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Enviar"]')))
                    botao_enviar.click()

                    time.sleep(1)

                    # Tenta localizar o XPath das oportunidades (e ignora se não achar)
                    try:
                        # 1. Volta para a URL salva
                        driver.get(url_ocorrencia)

                        # 2. Clica no botão de ocorrências
                        xpath_ocorrencia = '/html/body/div[2]/div[2]/ul/li[15]/a'
                        botao = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_ocorrencia)))
                        botao.click()

                        # 3. Aguarda carregar o corpo da tabela
                        xpath_tbody = '/html/body/div[2]/div[2]/div[2]/div/table/tbody'
                        tbody = wait.until(EC.presence_of_element_located((By.XPATH, xpath_tbody)))

                        # 4. Pega a primeira linha da tabela
                        linha = tbody.find_element(By.XPATH, './tr[1]')  # ou use tr diretamente no XPath anterior

                        # 5. Busca todos os <a> dentro da linha
                        links = linha.find_elements(By.TAG_NAME, 'a')

                        if len(links) >= 8:
                            # 6. Pega o href do 8º link
                            href_ocorrencia_criada = links[7].get_attribute('href')
                            print("Acessando link da ocorrência:", href_ocorrencia_criada)

                            # 7. Acessa o link da ocorrência criada
                            driver.get(href_ocorrencia_criada)
                        else:
                            print("Menos de 8 links encontrados na linha da tabela.")
                            
                        tbody = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[1]/div[2]/table/tbody')))
                        linhas = tbody.find_elements(By.TAG_NAME, 'tr')

                        oportunidades_disponiveis = []

                        for linha in linhas:
                            estilo = linha.get_attribute("style") or ""
                            possui_check_icon = bool(linha.find_elements(By.CLASS_NAME, "glyphicon-check"))

                            # Pula se o cliente já possui a oportunidade
                            if "#dff0d8" in estilo or possui_check_icon:
                                continue

                            colunas = linha.find_elements(By.TAG_NAME, 'td')
                            if colunas:
                                nome_oportunidade = colunas[0].text.strip()
                                if nome_oportunidade:
                                    oportunidades_disponiveis.append(nome_oportunidade)

                        if oportunidades_disponiveis:
                            nome_cliente = cliente.get("nome_cliente", "cliente")
                            mensagem = textwrap.dedent(f"""\
                                Olá {nome_cliente}, tudo bem?

                                Verifiquei aqui no nosso sistema e encontrei algumas oportunidades especiais disponíveis para você aproveitar agora mesmo:
                            """)

                            for item in oportunidades_disponiveis:
                                mensagem += f"- **{item}**\n"

                            mensagem += "\nSe quiser saber mais sobre alguma dessas oportunidades ou ativá-las agora, é só me chamar! Estou à disposição. 😊"

                            # Mostra a mensagem em um QDialog com campo para copiar
                            class DialogMensagem(QDialog):
                                def __init__(self, mensagem, parent=None):
                                    super().__init__(parent)
                                    self.setWindowTitle("Oportunidades para o cliente")
                                    layout = QVBoxLayout(self)

                                    label = QLabel("Mensagem pronta para enviar ao cliente:")
                                    layout.addWidget(label)

                                    self.text_edit = QTextEdit()
                                    self.text_edit.setPlainText(mensagem)
                                    self.text_edit.setReadOnly(True)
                                    layout.addWidget(self.text_edit)

                                    btn_copiar = QPushButton("Copiar para área de transferência")
                                    btn_copiar.clicked.connect(self.copiar_para_clipboard)
                                    layout.addWidget(btn_copiar)

                                def copiar_para_clipboard(self):
                                    clipboard = QApplication.clipboard()
                                    clipboard.setText(self.text_edit.toPlainText())
                                    QMessageBox.information(self, "Copiado", "Mensagem copiada com sucesso!")

                            dlg = DialogMensagem(mensagem)
                            dlg.exec_()

                        else:
                            print("Nenhuma nova oportunidade encontrada ou todas já ativadas.")

                    except TimeoutException:
                        # XPath não encontrado — apenas ignora
                        print("Oportunidades não disponíveis para esse cliente.")
                        
                    except Exception as e:
                        QMessageBox.warning(None, "Erro ao verificar oportunidades", f"Ocorreu um erro: {str(e)}")

                        # Envia
                        btn_salvar = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn-salvar"]')))
                        btn_salvar.click()

                    time.sleep(1)

                    # Tenta localizar e clicar no botão de fechar (se existir)
                    try:
                        xpath_fechar = '/html/body/div[2]/div[1]/div/a'
                        botao_fechar = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_fechar)))
                        botao_fechar.click()
                    except Exception as e:
                        print("Aviso: Não foi possível fechar a aba da ocorrência automaticamente.")

                except Exception as e:
                    QMessageBox.warning(janela_edicao, "Erro ao preencher ocorrência", f"Erro: {str(e)}")

            btn_ocorrencia.clicked.connect(gerar_ocorrencia)

            # Formulário
            form_layout = QFormLayout()
            inputs = {}

            campos_texto = [
                ("codigo", "Código"),
                ("nome_cliente", "Nome do Cliente"),
                ("email_contador", "E-mail do Contador"),
                ("email_secundario", "E-mail Secundário"),
                ("status", "Status"),
            ]

            for chave, rotulo in campos_texto:
                if chave == 'status':
                    entrada = QComboBox()
                    entrada.addItems(["BLOQUEADO", "CLIENTE GERA", "CONTADOR GERA", "NÃO GERA SPED","PENDENTE", "FEITO"])
                    atual = cliente.get(chave, "").upper()
                    if atual in ["BLOQUEADO", "CLIENTE GERA", "CONTADOR GERA", "NÃO GERA SPED","PENDENTE", "FEITO"]:
                        entrada.setCurrentText(atual)
                else:
                    entrada = QLineEdit(cliente.get(chave, ""))

                form_layout.addRow(QLabel(rotulo), entrada)
                inputs[chave] = entrada

            # Combos de pagamento
            combos_pagamento = {
                "pix_pdv": QComboBox(),
                "pix_off": QComboBox(),
                "pos_adiquirente": QComboBox(),
                "boleto": QComboBox(),
                "tef": QComboBox(),
                "delivery": QComboBox()
            }

            def populate_combo(combo, participantes):
                combo.addItem("Nenhum", "")
                for p in participantes:
                    display = f"{p.get('nome', 'Sem Nome')} - {p.get('nome_mun', 'N/A')}"
                    combo.addItem(display, p.get("codigo", ""))
                combo.setEditable(True)
                combo.setInsertPolicy(QComboBox.NoInsert)
                items = [combo.itemText(i) for i in range(combo.count())]
                completer = QCompleter(items, combo)
                completer.setCaseSensitivity(Qt.CaseInsensitive)
                completer.setFilterMode(Qt.MatchContains)
                combo.setCompleter(completer)

            participantes = getattr(self, "participantes", [])

            for chave, combo in combos_pagamento.items():
                populate_combo(combo, participantes)
                valor_atual = cliente.get(chave, "")
                index = combo.findData(valor_atual)
                combo.setCurrentIndex(index if index != -1 else 0)
                form_layout.addRow(QLabel(f"Banco para {chave.replace('_', ' ').upper()}"), combo)

            main_layout.addLayout(form_layout)

            # Botões Salvar/Cancelar
            btn_layout = QHBoxLayout()
            btn_salvar = QPushButton("Salvar Alterações")
            btn_cancelar = QPushButton("Cancelar")
            btn_layout.addWidget(btn_salvar)
            btn_layout.addWidget(btn_cancelar)
            main_layout.addLayout(btn_layout)

            # Função de salvar alterações
            def salvar_alteracoes():
                # 1. Atualiza os campos do cliente com os dados do formulário
                for chave, widget in inputs.items():
                    if isinstance(widget, QComboBox):
                        valor = widget.currentText().strip()
                    else:
                        valor = widget.text().strip()
                    cliente[chave] = valor

                for chave, combo in combos_pagamento.items():
                    cliente[chave] = combo.currentData()

                cliente["prioridade"] = "Sim" if checkbox_prioridade.isChecked() else "Não"

                # 2. Agora verifica se pode salvar direto ou deve salvar como pendência
                if not self.edicao_liberada:
                    self._salvar_pendencia(cliente)  # Salva já com as alterações!
                    self.atualizar_contador_pendencias()
                    QMessageBox.information(
                        self, "Modo somente leitura",
                        "A planilha está em uso por outro usuário.\nSuas alterações foram salvas como pendência."
                    )
                    janela_edicao.accept()
                    return

                # 3. Salva normalmente
                try:
                    self.salvar_clientes() 
                except Exception as e:
                    QMessageBox.warning(janela_edicao, "Erro", f"Erro ao salvar cliente: {e}")
                    return

                atualizar_tabela_clientes()
                janela_edicao.accept()

            btn_salvar.clicked.connect(salvar_alteracoes)
            btn_cancelar.clicked.connect(janela_edicao.reject)

            janela_edicao.exec()

        # Conecta o double-click
        self.tabela_clientes.itemDoubleClicked.connect(on_cliente_duplo_clique)

        # === Aba Participantes ===
        aba_participantes = QWidget()
        layout_participantes = QVBoxLayout()
        aba_participantes.setLayout(layout_participantes)

        # Container para botões em linha (horizontal)
        container_botoes = QWidget()
        botoes_layout = QHBoxLayout()
        botoes_layout.setContentsMargins(0, 0, 0, 0)
        botoes_layout.setSpacing(10)
        container_botoes.setLayout(botoes_layout)

        # Botão cadastrar participante
        btn_cadastrar = QPushButton("➕ Cadastrar Participante")
        btn_cadastrar.clicked.connect(self.cadastrar_participante)
        btn_cadastrar.setStyleSheet("font-size: 14px;")

        # Botão remover participante
        btn_remover = QPushButton("🗑️ Remover Participante")
        btn_remover.clicked.connect(self.remover_participante)
        btn_remover.setStyleSheet("font-size: 14px;")

        botoes_layout.addWidget(btn_cadastrar, alignment=Qt.AlignLeft)
        botoes_layout.addWidget(btn_remover, alignment=Qt.AlignLeft)
        botoes_layout.addStretch()

        layout_participantes.addWidget(container_botoes)

        # Campo de filtro
        filtro_participantes = QLineEdit()
        filtro_participantes.setPlaceholderText("Filtrar por nome ou cnpj")
        filtro_participantes.setFixedWidth(250)
        layout_participantes.addWidget(filtro_participantes)

        # Tabela de participantes
        self.tabela_participantes = QTableWidget()
        tabela_participantes = self.tabela_participantes
        tabela_participantes.setColumnCount(8)
        tabela_participantes.setHorizontalHeaderLabels([
            "Nome", "CNPJ", "Código do País", "Código do Município",
            "Nome do Município", "Logradouro", "Bairro", "Número"
        ])
        tabela_participantes.cellDoubleClicked.connect(
            lambda row, col: self.abrir_dados_participante(row, aba_participantes, ARQUIVO_PARTICIPANTES)
        )
        tabela_participantes.setEditTriggers(QTableWidget.NoEditTriggers)
        tabela_participantes.setSelectionBehavior(QTableWidget.SelectRows)
        tabela_participantes.setAlternatingRowColors(True)
        tabela_participantes.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)

        # Função para formatar CNPJ
        def formatar_cnpj(cnpj_raw):
            cnpj_num = ''.join(filter(str.isdigit, cnpj_raw))
            if len(cnpj_num) != 14:
                return cnpj_raw
            return f"{cnpj_num[:2]}.{cnpj_num[2:5]}.{cnpj_num[5:8]}/{cnpj_num[8:12]}-{cnpj_num[12:]}"

        # Preenche a tabela com os dados
        tabela_participantes.setRowCount(len(self.participantes))
        for i, p in enumerate(self.participantes):
            tabela_participantes.setItem(i, 0, QTableWidgetItem(p.get("nome", "")))
            cnpj_formatado = formatar_cnpj(p.get("cnpj", ""))
            tabela_participantes.setItem(i, 1, QTableWidgetItem(cnpj_formatado))
            cod_pais = p.get("cod_pais", "")
            if cod_pais == "1058":
                cod_pais = "1058 - Brasil"
            tabela_participantes.setItem(i, 2, QTableWidgetItem(cod_pais))
            tabela_participantes.setItem(i, 3, QTableWidgetItem(p.get("cod_mun", "")))
            tabela_participantes.setItem(i, 4, QTableWidgetItem(p.get("nome_mun", "")))
            tabela_participantes.setItem(i, 5, QTableWidgetItem(p.get("logradouro", "")))
            tabela_participantes.setItem(i, 6, QTableWidgetItem(p.get("bairro", "")))
            tabela_participantes.setItem(i, 7, QTableWidgetItem(p.get("SN", "")))

        tabela_participantes.horizontalHeader().setStretchLastSection(True)
        tabela_participantes.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # Função de filtro
        def aplicar_filtro_participantes(tabela, texto):
            texto = texto.lower()
            for row_idx in range(tabela.rowCount()):
                item_codigo = tabela.item(row_idx, 0)
                item_nome = tabela.item(row_idx, 1)
                codigo_texto = item_codigo.text().lower() if item_codigo else ""
                nome_texto = item_nome.text().lower() if item_nome else ""
                if not texto:
                    tabela.setRowHidden(row_idx, False)
                else:
                    if texto in codigo_texto or texto in nome_texto:
                        tabela.setRowHidden(row_idx, False)
                    else:
                        tabela.setRowHidden(row_idx, True)

        # Conecta o filtro ao campo de texto para filtrar conforme digita
        filtro_participantes.textChanged.connect(
            lambda t: aplicar_filtro_participantes(tabela_participantes, t.strip().lower())
        )

        layout_participantes.addWidget(tabela_participantes)
        tabs.addTab(aba_participantes, "Participantes")

        filtro_clientes.textChanged.connect(lambda texto: aplicar_filtro_participantes(self.tabela_clientes, texto))
        filtro_participantes.textChanged.connect(lambda texto: aplicar_filtro_participantes(tabela_participantes, texto))

        if not self.edicao_liberada:
            # Se não conseguiu o lock, mostra quem está usando e para aqui
            self.mostrar_dono_do_lock()
            
        else:
            # Se chegou aqui, lock foi adquirido e edição está liberada
            try:
                qtd_pendencias = self._contar_pendencias()
                print(">>> Pendências detectadas:", qtd_pendencias)
                if qtd_pendencias > 0:
                    QMessageBox.information(
                        self,
                        "Aviso de Pendências",
                        f"⚠️ Existem {qtd_pendencias} pendência(s) salvas.\nVocê pode visualizá-las no menu 'Visualizar Pendências'."
                    )
                    self._abrir_pendencias()
                    atualizar_tabela_clientes()
            except Exception as e:
                print("Erro ao verificar pendências:", e)

        def liberar_lock_ao_fechar():
            if self.edicao_liberada:
                self._liberar_lock()
            window.show()
        
        dialog.finished.connect(liberar_lock_ao_fechar)
        dialog.exec()

    def selecionar_arquivo_SPED(self):
        arquivos_SPED = [f for f in os.listdir() if f.endswith('.txt')]

        if not arquivos_SPED:
            QMessageBox.warning(self, "Aviso", "Nenhum arquivo SPED encontrado no diretório.")
            return None

        opcoes = []
        mapa_opcao_para_arquivo = {}

        for arquivo in arquivos_SPED:
            try:
                with open(arquivo, 'r', encoding='latin1') as f:
                    for linha in f:
                        if linha.startswith('|0000|'):
                            partes = linha.strip().split('|')
                            if len(partes) >= 8:
                                nome_empresa = partes[6]
                                cnpj = partes[7]
                                rotulo = f"{nome_empresa} ({cnpj})"
                                opcoes.append(rotulo)
                                mapa_opcao_para_arquivo[rotulo] = arquivo
                            break  # encontrou o |0000|, não precisa ler mais
            except Exception as e:
                print(f"Erro ao ler {arquivo}: {e}")

        if not opcoes:
            QMessageBox.warning(self, "Aviso", "Nenhum arquivo SPED válido encontrado.")
            return None

        opcao_selecionada, ok = QInputDialog.getItem(
            self,
            "Selecione Empresa",
            "Escolha a empresa (arquivo SPED):",
            opcoes,
            0,
            False
        )

        if ok and opcao_selecionada:
            arquivo_escolhido = mapa_opcao_para_arquivo[opcao_selecionada]
            QMessageBox.information(self, "Arquivo Selecionado", f"Você escolheu: {opcao_selecionada}")
            return arquivo_escolhido
        else:
            return None

    def salvar_clientes(self):
        """
        Regrava o arquivo clientes.csv com o conteúdo de self.clientes.
        Deve ser chamado após qualquer modificação em self.clientes.
        """
        caminho_csv = os.path.join(self.diretorio_csv, "clientes.csv")
        if not self.edicao_liberada:
            QMessageBox.warning(self, "Somente leitura", "A planilha está em uso por outro usuário. As alterações não foram salvas.")
            return
        try:
            # Se não houver clientes, ainda assim cria o CSV com cabeçalhos vazios
            fieldnames = list(self.clientes[0].keys()) if self.clientes else [
                "codigo", "nome_cliente", "email_contador", "email_secundario", "status",
                "pix_pdv", "pix_off", "pos_adiquirente", "boleto", "tef", "delivery", "prioridade"
            ]

            with open(caminho_csv, mode='w', newline='', encoding='latin1') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(self.clientes)

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao salvar clientes: {e}")

    def cadastrar_participante(self):
        dialog = self.FormCadastrarParticipante(self)
        if dialog.exec() == QDialog.Accepted:
            novo = dialog.novo_participante
            campos = ["codigo", "nome", "cod_pais", "cnpj", "cod_mun", "logradouro", "SN", "bairro", "endereco", "nome_mun"]
            # Se o arquivo não existir, cria e escreve o cabeçalho
            arquivo_existe = os.path.exists(self.arquivo_participantes)
            with open(self.arquivo_participantes, "a", newline='', encoding="latin1") as f:
                writer = csv.DictWriter(f, fieldnames=campos)
                if not arquivo_existe:
                    writer.writeheader()
                writer.writerow(novo)
            QMessageBox.information(self, "Sucesso", "Participante cadastrado com sucesso!")
            self.atualizar_tabela_participantes()
    
    def atualizar_tabela_participantes(self):
        if not os.path.exists(self.arquivo_participantes):
            return

        with open(self.arquivo_participantes, newline='', encoding="latin1") as f:
            leitor = csv.DictReader(f)
            dados = list(leitor)

        self.participantes = dados  # Atualiza a lista interna, se você usar em outro lugar

        self.tabela_participantes.setRowCount(len(dados))
        self.tabela_participantes.clearContents()

        def formatar_cnpj(cnpj_raw):
            cnpj_num = ''.join(filter(str.isdigit, cnpj_raw))
            if len(cnpj_num) != 14:
                return cnpj_raw
            return f"{cnpj_num[:2]}.{cnpj_num[2:5]}.{cnpj_num[5:8]}/{cnpj_num[8:12]}-{cnpj_num[12:]}"

        for i, p in enumerate(dados):
            self.tabela_participantes.setItem(i, 0, QTableWidgetItem(p.get("nome", "")))
            self.tabela_participantes.setItem(i, 1, QTableWidgetItem(formatar_cnpj(p.get("cnpj", ""))))
            
            cod_pais = p.get("cod_pais", "")
            if cod_pais == "1058":
                cod_pais = "1058 - Brasil"
            self.tabela_participantes.setItem(i, 2, QTableWidgetItem(cod_pais))
            self.tabela_participantes.setItem(i, 3, QTableWidgetItem(p.get("cod_mun", "")))
            self.tabela_participantes.setItem(i, 4, QTableWidgetItem(p.get("nome_mun", "")))
            self.tabela_participantes.setItem(i, 5, QTableWidgetItem(p.get("logradouro", "")))
            self.tabela_participantes.setItem(i, 6, QTableWidgetItem(p.get("bairro", "")))
            self.tabela_participantes.setItem(i, 7, QTableWidgetItem(p.get("SN", "")))

    def listar_participantes(self):
        """Exibe uma janela para seleção de um participante e, após a seleção, apresenta os dados completos para análise."""
        try:
            # Verifica se o arquivo de participantes existe
            if not os.path.exists(self.arquivo_participantes):
                QMessageBox.information(self, "Participantes", "Nenhum participante cadastrado.")
                return

            # Lê os participantes do arquivo CSV utilizando encoding 'latin1'
            with open(self.arquivo_participantes, newline='', encoding='latin1') as f:
                leitor = csv.DictReader(f)
                participantes = list(leitor)

            if not participantes:
                QMessageBox.information(self, "Participantes", "Nenhum participante cadastrado.")
                return

            # --- Diálogo para seleção do participante com filtro ---
            dialog_participante = QDialog(self)
            dialog_participante.setWindowTitle("Seleção de Participante")
            dialog_participante.setMinimumSize(400, 400)
            layout = QVBoxLayout(dialog_participante)

            # Campo de filtro para facilitar a busca
            filtro_line = QLineEdit(dialog_participante)
            filtro_line.setPlaceholderText("Filtrar participantes...")
            layout.addWidget(filtro_line)

            # Lista onde cada participante é exibido com informações resumidas
            list_widget = QListWidget(dialog_participante)
            for p in participantes:
                display_text = (
                    f"{p.get('nome', 'N/A')} - "
                    f"CNPJ: {p.get('codigo', 'N/A')} - "
                    f"Cidade: {p.get('nome_mun', 'N/A')}"
                )
                item = QListWidgetItem(display_text)
                # Armazena os dados completos do participante para uso posterior
                item.setData(Qt.UserRole, p)
                list_widget.addItem(item)
            layout.addWidget(list_widget)

            # Função que filtra os itens da lista conforme o texto digitado
            def filtrar_itens():
                texto = filtro_line.text().lower().strip()
                for i in range(list_widget.count()):
                    item = list_widget.item(i)
                    item.setHidden(texto not in item.text().lower())

            # Conecta o campo de texto à função de filtro
            filtro_line.textChanged.connect(filtrar_itens)

            # Botões de OK e Cancelar
            btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            layout.addWidget(btn_box)
            btn_box.accepted.connect(dialog_participante.accept)
            btn_box.rejected.connect(dialog_participante.reject)

            # Executa o diálogo para a seleção
            if dialog_participante.exec() == QDialog.Accepted:
                selected_item = list_widget.currentItem()
                if not selected_item:
                    QMessageBox.information(self, "Informação", "Nenhum participante selecionado.")
                    return
                participante_selecionado = selected_item.data(Qt.UserRole)

                # --- Exibe os dados completos do participante selecionado ---
                dialog_dados = QDialog(self)
                dialog_dados.setWindowTitle("Detalhes do Participante")
                dialog_dados.setMinimumSize(400, 400)
                layout_dados = QVBoxLayout(dialog_dados)
                text_edit = QTextEdit(dialog_dados)
                text_edit.setReadOnly(True)

                # Formata os dados para exibição com ordem personalizada
                dados_texto = "=== Dados do Participante ===\n\n"
                # Lista com a ordem desejada para exibição. Altere para os campos que desejar.
                ordem = ['nome', 'codigo', 'cod_pais', 'cod_mun', 'nome_mun', 'logradouro', 'bairro', 'SN']
                for chave in ordem:
                    valor = participante_selecionado.get(chave, 'N/A')
                    dados_texto += f"{chave}: {valor}\n"

                text_edit.setPlainText(dados_texto)
                layout_dados.addWidget(text_edit)

                # Botão para fechar a janela de detalhes
                btn_fechar = QPushButton("Fechar", dialog_dados)
                btn_fechar.clicked.connect(dialog_dados.accept)
                layout_dados.addWidget(btn_fechar)

                dialog_dados.exec()
            else:
                return

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao carregar participantes: {e}")
    
    def cadastrar_cliente(self):
        dialog = self.FormCadastrarCliente(self)
        if dialog.exec() == QDialog.Accepted:
            novo_cliente = dialog.novo_cliente
            path_clientes = os.path.join(self.diretorio_csv, "clientes.csv")
            fieldnames = [
                'codigo', 'nome_cliente', 'email_contador', 'email_secundario', 'status',
                'pix_pdv', 'pix_off', 'pos_adiquirente', 'boleto', 'tef', 'delivery'
            ]

            usuario_atual = getpass.getuser()
            maquina_atual = socket.gethostname()

            lock_info = self._ler_lock_info()
            lock_existe = lock_info is not None
            lock_e_meu = lock_existe and (lock_info.get("usuario") == usuario_atual and lock_info.get("maquina") == maquina_atual)

            # Se existe lock e não é seu, salva pendência
            if lock_existe and not lock_e_meu:
                novo_cliente["prioridade"] = ""
                novo_cliente["status"] = "PENDENTE"
                self._salvar_pendencia(novo_cliente)
                QMessageBox.information(
                    self,
                    "Pendente",
                    "Cadastro salvo como pendência.\nImporte quando a planilha estiver liberada."
                )
                return

            # Se não existe lock (algo estranho, mas vamos salvar como pendência)
            if not lock_existe:
                novo_cliente["prioridade"] = ""
                novo_cliente["status"] = "PENDENTE"
                self._salvar_pendencia(novo_cliente)
                QMessageBox.information(
                    self,
                    "Pendente",
                    "Cadastro salvo como pendência.\nImporte quando a planilha estiver liberada."
                )
                return

            # Se chegou aqui, lock existe e é seu — pode cadastrar direto
            try:
                novo_arquivo = not os.path.exists(path_clientes)
                with open(path_clientes, mode='a', newline='', encoding='latin1') as f:
                    writer = csv.DictWriter(f, fieldnames=fieldnames)
                    if novo_arquivo:
                        writer.writeheader()
                    writer.writerow(novo_cliente)
                QMessageBox.information(self, "Sucesso", "Cliente cadastrado com sucesso!")
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao cadastrar cliente: {e}")

    def remover_cliente(self):
        """
        Remove um cliente selecionado e regrava o CSV.
        Só remove se o lock for seu (ou seja, você tem permissão para editar).
        """

        # Verifica se há clientes em memória
        if not hasattr(self, "clientes") or not self.clientes:
            QMessageBox.information(self, "Remover Cliente", "Nenhum cliente cadastrado.")
            return

        usuario_atual = getpass.getuser()
        maquina_atual = socket.gethostname()
        lock_info = self._ler_lock_info()
        lock_existe = lock_info is not None
        lock_e_meu = lock_existe and (lock_info.get("usuario") == usuario_atual and lock_info.get("maquina") == maquina_atual)

        if not lock_e_meu:
            QMessageBox.warning(
                self,
                "Planilha bloqueada",
                "A planilha está em uso por outro usuário.\nRemoção não permitida no momento."
            )
            return

        self.hide()
        # Seleção gráfica do cliente
        dialog = FormSelecionarCliente(self.clientes, self)
        if dialog.exec() != QDialog.Accepted:
            return  # usuário cancelou
        cliente = dialog.selected_client

        # Confirmação
        resp = QMessageBox.question(
            self,
            "Confirmar Remoção",
            f"Tem certeza que deseja remover o cliente:\n\n"
            f"{cliente.get('nome_cliente','')} (Código: {cliente.get('codigo','')})?",
            QMessageBox.Yes | QMessageBox.No
        )
        if resp != QMessageBox.Yes:
            return

        # Remove da lista em memória
        self.clientes = [c for c in self.clientes if c.get("codigo","") != cliente.get("codigo","")]

        # Grava no CSV
        try:
            self.salvar_clientes()
            QMessageBox.information(self, "Sucesso", "Cliente removido com sucesso!")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao remover cliente: {e}")


    def enviar_email_contador_com_cliente(self, cliente):
        pasta_base = DIRETORIO_SPED
        json_path = os.path.join(pasta_base, "config_email.json")

        def dialog_configurar_email():
            class EmailConfigDialog(QDialog):
                def __init__(self, parent=None):
                    super().__init__(parent)
                    self.setWindowTitle("Configuração de Envio de E-mail")
                    layout = QVBoxLayout()

                    self.nome_input = QLineEdit()
                    self.email_input = QLineEdit()
                    self.senha_input = QLineEdit()
                    self.senha_input.setEchoMode(QLineEdit.Password)
                    self.mensagem_input = QPlainTextEdit()

                    layout.addWidget(QLabel("Nome Identificador:"))
                    layout.addWidget(self.nome_input)
                    layout.addWidget(QLabel("E-mail remetente:"))
                    layout.addWidget(self.email_input)
                    layout.addWidget(QLabel("Chave de acesso (senha do app):"))
                    layout.addWidget(self.senha_input)
                    layout.addWidget(QLabel("Mensagem padrão (corpo do e-mail):"))
                    layout.addWidget(self.mensagem_input)

                    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
                    buttons.accepted.connect(self.accept)
                    buttons.rejected.connect(self.reject)
                    layout.addWidget(buttons)

                    self.setLayout(layout)

            dialog = EmailConfigDialog(self)
            if dialog.exec() != QDialog.Accepted:
                return None

            nome = dialog.nome_input.text().strip()
            email = dialog.email_input.text().strip()
            senha = dialog.senha_input.text().strip()
            mensagem = dialog.mensagem_input.toPlainText()

            if not nome or not email or not senha or not mensagem:
                QMessageBox.warning(self, "Campos Incompletos", "Todos os campos devem ser preenchidos.")
                return None

            return nome, email, senha, mensagem

        # --------- Nova caixa de seleção com botão Novo separado ---------

        config_data = {}

        if os.path.exists(json_path):
            try:
                with open(json_path, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao ler arquivo de configuração:\n{e}")
                return

            items = list(config_data.keys())

            dialog = QDialog(self)
            dialog.setWindowTitle("Selecionar Mensagem")

            layout = QVBoxLayout(dialog)
            combo = QComboBox()
            combo.addItems(items)
            layout.addWidget(QLabel("Escolha a mensagem para enviar:"))
            layout.addWidget(combo)

            button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            novo_button = QPushButton("Novo")

            button_box.addButton(novo_button, QDialogButtonBox.ActionRole)
            layout.addWidget(button_box)

            selected_item = None

            def on_ok():
                nonlocal selected_item
                selected_item = combo.currentText()
                dialog.accept()

            def on_novo():
                nova = dialog_configurar_email()
                if nova is None:
                    return
                nome, email, senha, mensagem = nova
                config_data[nome] = {
                    "email": email,
                    "senha": senha,
                    "mensagem": mensagem
                }
                try:
                    with open(json_path, 'w', encoding='utf-8') as f:
                        json.dump(config_data, f, indent=4, ensure_ascii=False)
                except Exception as e:
                    QMessageBox.critical(self, "Erro ao salvar configuração", str(e))
                    return
                combo.addItem(nome)
                combo.setCurrentText(nome)

            button_box.accepted.connect(on_ok)
            button_box.rejected.connect(dialog.reject)
            novo_button.clicked.connect(on_novo)

            if dialog.exec() != QDialog.Accepted:
                return

            remetente_info = config_data[selected_item]

        else:
            nova = dialog_configurar_email()
            if nova is None:
                return
            nome, email, senha, mensagem = nova

            config_data[nome] = {
                "email": email,
                "senha": senha,
                "mensagem": mensagem
            }

            try:
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(config_data, f, indent=4, ensure_ascii=False)
            except Exception as e:
                QMessageBox.critical(self, "Erro ao salvar configuração", str(e))
                return
            remetente_info = config_data[nome]

        # --------- Dados para envio ---------

        remetente = remetente_info.get("email")
        senha_app = remetente_info.get("senha")
        corpo = remetente_info.get("mensagem")

        emails = [e.strip() for e in (cliente.get('email_contador', ''), cliente.get('email_secundario', '')) if e.strip()]
        if not emails:
            QMessageBox.warning(self, "Sem E-mails", "Não há e-mails cadastrados para este cliente.")
            return

        filtro = "Arquivos TXT (*.txt)"
        caminho_SPED, _ = QFileDialog.getOpenFileName(self, "Selecione o arquivo SPED (.txt)", "", filtro)
        if not caminho_SPED:
            return

        nome_empresa = cnpj_empresa = ""
        try:
            with open(caminho_SPED, 'r', encoding='latin1') as f:
                for linha in f:
                    if linha.startswith("|0000|"):
                        partes = linha.strip().split("|")
                        if len(partes) >= 8:
                            nome_empresa, cnpj_empresa = partes[6].strip(), partes[7].strip()
                        break
        except Exception as e:
            QMessageBox.critical(self, "Erro Leitura SPED", f"Não foi possível ler o SPED:\n{e}")
            return

        if not nome_empresa or not cnpj_empresa:
            QMessageBox.warning(self, "Dados Inválidos", "Não foi possível extrair empresa/CNPJ do SPED.")
            return

        hoje = datetime.now()
        mes = hoje.month - 1 or 12
        ano = hoje.year if hoje.month != 1 else hoje.year - 1
        mes_str = f"{mes:02d}"
        assunto = f"SPED {mes_str}/{ano} - {nome_empresa} - {cnpj_empresa}"

        dominios_comuns = ['gmail.com', 'hotmail.com', 'outlook.com']
        automaticos, manuais = [], []
        for e in emails:
            dom = e.split('@')[-1].lower()
            (automaticos if dom in dominios_comuns else manuais).append(e)

        if automaticos:
            try:
                msg = EmailMessage()
                msg['From'] = remetente
                msg['To'] = ', '.join(automaticos)
                msg['Subject'] = assunto
                msg.set_content(corpo)
                with open(caminho_SPED, 'rb') as f:
                    msg.add_attachment(f.read(),
                                    maintype='application',
                                    subtype='octet-stream',
                                    filename=os.path.basename(caminho_SPED))
                with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                    smtp.login(remetente, senha_app)
                    smtp.send_message(msg)
                QMessageBox.information(self, "E-mail Enviado",
                                        f"E-mail enviado com sucesso para:\n{', '.join(automaticos)}")
            except Exception as e:
                QMessageBox.critical(self, "Erro Envio SMTP", f"Falha ao enviar e-mail:\n{e}")
                return

        if manuais:
            to_param = urllib.parse.quote(','.join(manuais))
            assunto_param = urllib.parse.quote(assunto)
            corpo_param = urllib.parse.quote(corpo)

            url = (f"https://mail.google.com/mail/u/0/?view=cm&fs=1"
                f"&to={to_param}&su={assunto_param}&body={corpo_param}")
            webbrowser.open(url)
            QMessageBox.information(self, "Envio Manual",
                                    f"Abra o navegador para completar o envio para:\n{', '.join(manuais)}")

        time.sleep(1)

    def enviar_email_contador(self):
        pasta_base = DIRETORIO_SPED
        json_path = os.path.join(pasta_base, "config_email.json")

        def dialog_configurar_email():
            class EmailConfigDialog(QDialog):
                def __init__(self, parent=None):
                    super().__init__(parent)
                    self.setWindowTitle("Configuração de Envio de E-mail")
                    layout = QVBoxLayout()

                    self.nome_input = QLineEdit()
                    self.email_input = QLineEdit()
                    self.senha_input = QLineEdit()
                    self.senha_input.setEchoMode(QLineEdit.Password)
                    self.mensagem_input = QPlainTextEdit()

                    layout.addWidget(QLabel("Nome Identificador:"))
                    layout.addWidget(self.nome_input)
                    layout.addWidget(QLabel("E-mail remetente:"))
                    layout.addWidget(self.email_input)
                    layout.addWidget(QLabel("Chave de acesso (senha do app):"))
                    layout.addWidget(self.senha_input)
                    layout.addWidget(QLabel("Mensagem padrão (corpo do e-mail):"))
                    layout.addWidget(self.mensagem_input)

                    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
                    buttons.accepted.connect(self.accept)
                    buttons.rejected.connect(self.reject)
                    layout.addWidget(buttons)

                    self.setLayout(layout)

            dialog = EmailConfigDialog(self)
            if dialog.exec() != QDialog.Accepted:
                return None

            nome = dialog.nome_input.text().strip()
            email = dialog.email_input.text().strip()
            senha = dialog.senha_input.text().strip()
            mensagem = dialog.mensagem_input.toPlainText()

            if not nome or not email or not senha or not mensagem:
                QMessageBox.warning(self, "Campos Incompletos", "Todos os campos devem ser preenchidos.")
                return None

            return nome, email, senha, mensagem

        # ---------- Carregar ou criar configuração ----------
        config_data = {}
        if os.path.exists(json_path):
            try:
                with open(json_path, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao ler arquivo de configuração:\n{e}")
                return
        else:
            # Se o arquivo não existe, cria o primeiro remetente
            nova = dialog_configurar_email()
            if nova is None:
                return
            nome, email, senha, mensagem = nova
            config_data[nome] = {
                "email": email,
                "senha": senha,
                "mensagem": mensagem
            }
            try:
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(config_data, f, indent=4, ensure_ascii=False)
            except Exception as e:
                QMessageBox.critical(self, "Erro ao salvar configuração", str(e))
                return

        # Interface de seleção com botão Novo
        class EscolherMensagemDialog(QDialog):
            def __init__(self, config_data, parent=None):
                super().__init__(parent)
                self.setWindowTitle("Selecionar Mensagem")
                self.resultado = None

                layout = QVBoxLayout(self)
                self.combo = QComboBox(self)
                self.combo.addItems(config_data.keys())
                layout.addWidget(self.combo)

                button_box = QHBoxLayout()
                self.ok_button = QPushButton("Ok", self)
                self.cancel_button = QPushButton("Cancelar", self)
                self.novo_button = QPushButton("Novo", self)

                self.ok_button.clicked.connect(self.accept)
                self.cancel_button.clicked.connect(self.reject)
                self.novo_button.clicked.connect(self.novo)

                button_box.addWidget(self.ok_button)
                button_box.addWidget(self.novo_button)
                button_box.addWidget(self.cancel_button)

                layout.addLayout(button_box)

            def novo(self):
                self.resultado = "NOVO"
                self.done(2)  # Código diferente para o Novo

        dialog = EscolherMensagemDialog(config_data, self)
        res = dialog.exec()

        if res == 0:  # Cancelado
            return
        elif res == 2:  # Novo pressionado
            nova = dialog_configurar_email()
            if nova is None:
                return
            nome, email, senha, mensagem = nova
            config_data[nome] = {
                "email": email,
                "senha": senha,
                "mensagem": mensagem
            }
            try:
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(config_data, f, indent=4, ensure_ascii=False)
            except Exception as e:
                QMessageBox.critical(self, "Erro ao salvar configuração", str(e))
                return
            remetente_info = config_data[nome]
        else:
            nome_selecionado = dialog.combo.currentText()
            remetente_info = config_data[nome_selecionado]

        remetente = remetente_info.get("email")
        senha_app = remetente_info.get("senha")
        corpo = remetente_info.get("mensagem")

        # ---------- Selecionar cliente ----------
        self.carregar_clientes(ARQUIVO_CLIENTES)
        dialog_cliente = FormSelecionarCliente(self.clientes, self)
        if dialog_cliente.exec() != QDialog.Accepted or not dialog_cliente.selected_client:
            return
        cliente = dialog_cliente.selected_client

        emails = [e.strip() for e in (cliente.get('email_contador', ''), cliente.get('email_secundario', '')) if e.strip()]
        if not emails:
            QMessageBox.warning(self, "Sem E-mails", "Não há e-mails cadastrados para este cliente.")
            return

        # ---------- Selecionar SPED ----------
        filtro = "Arquivos TXT (*.txt)"
        caminho_SPED, _ = QFileDialog.getOpenFileName(self, "Selecione o arquivo SPED (.txt)", "", filtro)
        if not caminho_SPED:
            return

        nome_empresa = cnpj_empresa = ""
        try:
            with open(caminho_SPED, 'r', encoding='latin1') as f:
                for linha in f:
                    if linha.startswith("|0000|"):
                        partes = linha.strip().split("|")
                        if len(partes) >= 8:
                            nome_empresa, cnpj_empresa = partes[6].strip(), partes[7].strip()
                        break
        except Exception as e:
            QMessageBox.critical(self, "Erro Leitura SPED", f"Não foi possível ler o SPED:\n{e}")
            return

        if not nome_empresa or not cnpj_empresa:
            QMessageBox.warning(self, "Dados Inválidos", "Não foi possível extrair empresa/CNPJ do SPED.")
            return

        hoje = datetime.now()
        mes = hoje.month - 1 or 12
        ano = hoje.year if hoje.month != 1 else hoje.year - 1
        mes_str = f"{mes:02d}"
        assunto = f"SPED {mes_str}/{ano} - {nome_empresa} - {cnpj_empresa}"

        dominios_comuns = ['gmail.com', 'hotmail.com', 'outlook.com']
        automaticos, manuais = [], []
        for e in emails:
            dom = e.split('@')[-1].lower()
            (automaticos if dom in dominios_comuns else manuais).append(e)

        if automaticos:
            try:
                msg = EmailMessage()
                msg['From'] = remetente
                msg['To'] = ', '.join(automaticos)
                msg['Subject'] = assunto
                msg.set_content(corpo)
                with open(caminho_SPED, 'rb') as f:
                    msg.add_attachment(f.read(),
                                    maintype='application',
                                    subtype='octet-stream',
                                    filename=os.path.basename(caminho_SPED))
                with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                    smtp.login(remetente, senha_app)
                    smtp.send_message(msg)
                QMessageBox.information(self, "E-mail Enviado",
                                        f"E-mail enviado com sucesso para:\n{', '.join(automaticos)}")
            except Exception as e:
                QMessageBox.critical(self, "Erro Envio SMTP", f"Falha ao enviar e-mail:\n{e}")
                return

        if manuais:
            to_param = urllib.parse.quote(','.join(manuais))
            assunto_param = urllib.parse.quote(assunto)
            corpo_param = urllib.parse.quote(corpo)

            url = (f"https://mail.google.com/mail/u/0/?view=cm&fs=1"
                f"&to={to_param}&su={assunto_param}&body={corpo_param}")
            webbrowser.open(url)
            QMessageBox.information(self, "Envio Manual",
                                    f"Abra o navegador para completar o envio para:\n{', '.join(manuais)}")
        time.sleep(1)
        
    def fechar_aplicacao(self):
        self.close()

    def trocar_banco_cliente(self):
        usuario_atual = getpass.getuser()
        maquina_atual = socket.gethostname()

        lock_info = self._ler_lock_info()
        lock_existe = lock_info is not None
        lock_e_meu = lock_existe and (lock_info.get("usuario") == usuario_atual and lock_info.get("maquina") == maquina_atual)

        if not lock_existe:
            QMessageBox.warning(self, "Acesso Negado", "A planilha não está bloqueada para edição.\nNão é possível alterar os dados.")
            return
        if not lock_e_meu:
            QMessageBox.warning(self, "Acesso Negado", "A planilha está sendo usada por outro usuário.\nAlteração não permitida.")
            return

        # Continua com o código original, pois só chega aqui se for o dono do lock
        self.carregar_clientes(os.path.join(self.diretorio_csv, 'clientes.csv'))

        participantes = self.carregar_participantes()
        if not participantes:
            QMessageBox.information(self, "Informação", "Nenhum participante cadastrado.")
            return

        # Seleção do Cliente
        dialog_cliente = FormSelecionarCliente(self.clientes, self)
        if dialog_cliente.exec() == QDialog.Accepted:
            selected_item = dialog_cliente.list_widget.currentItem()
            if not selected_item:
                QMessageBox.information(self, "Informação", "Nenhum cliente selecionado.")
                return
            cliente_selecionado = selected_item.data(Qt.UserRole)
        else:
            return

        # Escolha da Ação
        opcoes = ["E-mails do cliente", "Bancos do cliente", "Status do cliente"]
        escolha_acao, ok = QInputDialog.getItem(
            self, "Trocar Dados do Cliente",
            "O que deseja alterar?",
            opcoes, 0, False
        )
        if not ok:
            return

        # Alterar Status do cliente
        if escolha_acao == "Status do cliente":
            atual = cliente_selecionado.get("status", "")
            novo, ok = QInputDialog.getItem(
                self, "Alterar Status",
                f"Status atual: '{atual or '(vazio)'}'. Marcar como:",
                ["FEITO", "(deixar vazio)"], 0, False
            )
            if ok:
                cliente_selecionado["status"] = "" if novo == "(deixar vazio)" else "FEITO"
                QMessageBox.information(self, "Sucesso", f"Status alterado para '{cliente_selecionado['status'] or '(vazio)'}'.")
            else:
                return

        # Alterar E-mails do cliente
        elif escolha_acao == "E-mails do cliente":
            while True:
                opcoes_email = [
                    "1. E-mail do contador: " + cliente_selecionado.get("email_contador", ""),
                    "2. E-mail secundário: " + cliente_selecionado.get("email_secundario", "")
                ]
                opcao_email, ok = QInputDialog.getItem(
                    self, "Alterar E-mails",
                    "Selecione o e-mail que deseja alterar (ou cancele para sair):",
                    opcoes_email, 0, False
                )
                if not ok:
                    break
                campo = "email_contador" if opcao_email.startswith("1.") else "email_secundario"
                acao, ok = QInputDialog.getItem(
                    self, "Alterar E-mail",
                    "Deseja Substituir ou Remover este e-mail?",
                    ["Substituir", "Remover"], 0, False
                )
                if not ok:
                    continue
                if acao == "Substituir":
                    novo_email, ok = QInputDialog.getText(
                        self, "Novo E-mail", "Digite o novo e-mail:"
                    )
                    if ok:
                        cliente_selecionado[campo] = novo_email
                        QMessageBox.information(self, "Sucesso", "E-mail atualizado com sucesso.")
                else:
                    cliente_selecionado[campo] = ""
                    QMessageBox.information(self, "Sucesso", "E-mail removido com sucesso.")

        # Alterar Bancos do cliente
        elif escolha_acao == "Bancos do cliente":
            formas = ["PIX PDV", "PIX OFF", "POS ADIQUIRENTE", "BOLETO", "TEF", "DELIVERY"]
            lista_metodos = []
            for metodo in formas:
                chave_metodo = metodo.lower().replace(" ", "_")
                codigo = cliente_selecionado.get(chave_metodo)
                if codigo:
                    info = next((p for p in participantes if p['codigo'] == codigo), None)
                    if info:
                        lista_metodos.append(f"{metodo}: {info.get('nome')} (cnpj:{codigo})")
                    else:
                        lista_metodos.append(f"{metodo}: Banco não encontrado")
                else:
                    lista_metodos.append(f"{metodo}:")
            escolha_str, ok = QInputDialog.getItem(
                self, "Alterar Bancos",
                "Selecione a forma de pagamento:",
                lista_metodos, 0, False
            )
            if not ok:
                return
            idx = lista_metodos.index(escolha_str)
            chave_metodo = formas[idx].lower().replace(" ", "_")
            acao_banco, ok = QInputDialog.getItem(
                self, "Alterar Banco",
                "Substituir ou Remover?",
                ["Substituir", "Remover"], 0, False
            )
            if not ok:
                return
            if acao_banco == "Remover":
                cliente_selecionado[chave_metodo] = ""
                QMessageBox.information(self, "Sucesso", f"Banco removido de {formas[idx]}.")
            else:
                dialog_banco = FormSelecionarBanco(participantes, self)
                if dialog_banco.exec() == QDialog.Accepted:
                    sel = dialog_banco.list_widget.currentItem()
                    if not sel:
                        QMessageBox.warning(self, "Erro", "Nenhum banco selecionado.")
                        return
                    banco = sel.data(Qt.UserRole)
                    cliente_selecionado[chave_metodo] = banco['codigo']
                    QMessageBox.information(self, "Sucesso", f"Banco {banco['nome']} associado.")
        else:
            QMessageBox.information(self, "Informação", "Opção inválida.")
            return

        # Atualiza lista e salva no CSV
        for i, cli in enumerate(self.clientes):
            if cli.get("codigo") == cliente_selecionado.get("codigo"):
                self.clientes[i] = cliente_selecionado
                break
        self.salvar_clientes()
        QMessageBox.information(self, "Sucesso", "Alterações salvas com sucesso.")


    def show_text_dialog(self, title, text):
        """Exibe um diálogo com o título e o texto fornecidos."""
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.setMinimumSize(600, 400)
        layout = QVBoxLayout(dialog)
        text_edit = QTextEdit(dialog)
        text_edit.setReadOnly(True)
        text_edit.setPlainText(text)
        layout.addWidget(text_edit)
        button_box = QDialogButtonBox(QDialogButtonBox.Ok)
        button_box.accepted.connect(dialog.accept)
        layout.addWidget(button_box)
        dialog.exec()

    def limpar_registros_1601(self):
        arquivo_txt = self.selecionar_arquivo_SPED()
        
        if not arquivo_txt:
            QMessageBox.information(self, "Operação cancelada", "Operação cancelada.")
            return

        try:
            caminho_base_participantes = os.path.join(DIRETORIO_SPED, "participantes.csv")
            base = pd.read_csv(caminho_base_participantes, dtype=str, encoding='latin1').fillna("")
            base["cnpj"] = base["cnpj"].apply(lambda c: re.sub(r'\D', '', c))

            with open(arquivo_txt, 'r', encoding='latin1') as f:
                linhas = f.readlines()

            registros_1601 = [l for l in linhas if l.startswith("|1601|")]
            if not registros_1601:
                QMessageBox.information(self, "Aviso", "Nenhum registro 1601 encontrado no arquivo.")
                return

            cnpjs_1601 = [re.sub(r'\D', '', l.split("|")[2]) for l in registros_1601 if len(l.split("|")) > 2]

            chaves_0150_a_remover = set()
            for cnpj in cnpjs_1601:
                linha_base = base[base["cnpj"] == cnpj]
                if not linha_base.empty:
                    p = linha_base.iloc[0]
                    chave = f"|0150|{cnpj}|{p['cod_pais']}|{cnpj}"
                    chaves_0150_a_remover.add(chave)

            novas_linhas = []
            removidos_0150 = 0
            for linha in linhas:
                if linha.startswith("|1601|"):
                    continue
                if linha.startswith("|0150|"):
                    campos = linha.strip().split("|")
                    if len(campos) >= 6:
                        chave_linha = f"|0150|{campos[2]}|{campos[4]}|{campos[5]}"
                        if chave_linha in chaves_0150_a_remover:
                            removidos_0150 += 1
                            continue
                novas_linhas.append(linha)

            qtd_removidos_1601 = len(registros_1601)

            self.atualizar_bloco_9900(novas_linhas, "1601", -qtd_removidos_1601)
            self.atualizar_bloco_9900(novas_linhas, "0150", -removidos_0150)
            self.atualizar_bloco_1990(novas_linhas, "1601")
            self.atualizar_bloco_9999(novas_linhas, sum(1 for l in novas_linhas if l.strip()))
            self.atualizar_bloco_0990(novas_linhas, -removidos_0150)

            with open(arquivo_txt, 'w', encoding='latin1') as f:
                f.writelines(novas_linhas)

            QMessageBox.information(self, "Limpeza concluída",
                                    f"{qtd_removidos_1601} registros 1601 removidos com sucesso.\n"
                                    f"{removidos_0150} blocos 0150 removidos também.")

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao limpar registros 1601 e 0150: {e}")


    def remover_participante(self):
        # Verifica se há participantes cadastrados
        if not self.participantes:
            QMessageBox.information(self, "Remover Participante", "Nenhum participante cadastrado.")
            return

        # Cria uma lista formatada com os participantes para exibição
        lista_participantes = [
            f"{i+1}. {p['nome']} (CNPJ: {p['cnpj']})" for i, p in enumerate(self.participantes)
        ]
        
        # Abre um diálogo para o usuário selecionar o participante a ser removido
        item, ok = QInputDialog.getItem(
            self,
            "Remover Participante",
            "Selecione o participante que deseja remover:",
            lista_participantes,
            0,
            False
        )
        
        if not ok or not item:
            return  # Operação cancelada

        # Tenta extrair o índice a partir da string selecionada
        try:
            idx = int(item.split('.')[0]) - 1  # O número vem antes do ponto
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao interpretar a seleção: {e}")
            return

        # Verifica se o índice é válido
        if idx < 0 or idx >= len(self.participantes):
            QMessageBox.critical(self, "Erro", "Seleção inválida.")
            return

        # Remove o participante da lista
        participante_removido = self.participantes.pop(idx)

        # Diretório SPED dos CSVs – adapte se necessário
        caminho_participantes = os.path.join(DIRETORIO_SPED, "participantes.csv")

        try:
            with open(caminho_participantes, 'w', newline='', encoding='latin1') as f:
                writer = csv.DictWriter(f, fieldnames=[
                    "codigo", "nome", "cod_pais", "cnpj", "cod_mun",
                    "logradouro", "SN", "bairro", "endereco", "nome_mun"
                ])
                writer.writeheader()
                writer.writerows(self.participantes)
            
            QMessageBox.information(
                self,
                "Sucesso",
                f"Participante '{participante_removido['nome']}' removido com sucesso."
            )
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao remover participante: {e}")

    def inserir_registro_1601_com_cliente(self, cliente_selecionado, atualizar_tabela_clientes):
        try:
            # Seleciona o arquivo SPED
            arquivo_txt = self.selecionar_arquivo_SPED()
            if not arquivo_txt:
                input("\nOperação cancelada. Pressione Enter para continuar...")
                return

            self.arquivo_SPED = arquivo_txt

            # Carregar participantes
            caminho_participantes = os.path.join(DIRETORIO_SPED, "participantes.csv")
            try:
                with open(caminho_participantes, newline='', encoding='latin1') as f:
                    leitor = csv.DictReader(f)
                    self.participantes = list(leitor)
            except Exception as e:
                print(f"\nErro ao carregar participantes: {e}")
                input("Pressione Enter para continuar...")
                return

            # --- Após selecionar o cliente, começa inserção dos 1601 ---
            cnpjs_ja_inseridos = []

            while True:
                meios_pagamento = ["pix_pdv", "pix_off", "pos_adiquirente", "boleto", "tef", "delivery"]
                bancos_disponiveis = {}
                indice_global = 1

                nao_registrados = []
                ja_registrados = []

                for metodo in meios_pagamento:
                    cnpj_banco = cliente_selecionado.get(metodo)
                    if cnpj_banco:
                        participante = next((p for p in self.participantes if p['cnpj'] == cnpj_banco), None)
                        if participante:
                            status = "JÁ INSERIDO" if cnpj_banco in cnpjs_ja_inseridos else "DISPONÍVEL"
                            entrada = {
                                "metodo": metodo,
                                "cnpj": cnpj_banco,
                                "nome": participante['nome'],
                                "municipio": participante['nome_mun'],
                                "participante": participante,
                                "status": status
                            }
                            if status == "DISPONÍVEL":
                                nao_registrados.append(entrada)
                            else:
                                ja_registrados.append(entrada)

                bancos_disponiveis = {}
                indice_global = 1
                lista_bancos = []

                if nao_registrados:
                    lista_bancos.append("--- Ainda não registrados ---")
                    for b in nao_registrados:
                        linha = f"{indice_global}. {b['metodo'].upper()}: {b['nome']} - CNPJ: {b['cnpj']} - Município: {b['municipio']} [DISPONÍVEL]"
                        lista_bancos.append(linha)
                        bancos_disponiveis[indice_global] = b
                        indice_global += 1

                if ja_registrados:
                    lista_bancos.append("--- Já inseridos anteriormente ---")
                    for b in ja_registrados:
                        linha = f"{indice_global}. {b['metodo'].upper()}: {b['nome']} - CNPJ: {b['cnpj']} - Município: {b['municipio']} [JÁ INSERIDO]"
                        lista_bancos.append(linha)
                        bancos_disponiveis[indice_global] = b
                        indice_global += 1

                # Caixa para seleção do banco
                itens_selecionaveis = [bancos_disponiveis[i]['metodo'].upper() + " - " + bancos_disponiveis[i]['nome'] for i in bancos_disponiveis]
                item, ok = QInputDialog.getItem(None, "Selecionar Banco", "Escolha o banco para inserir/somar valor:", itens_selecionaveis, 0, False)
                if not ok or not item:
                    QMessageBox.information(None, "Aviso", "Operação cancelada na seleção do banco.")
                    return

                banco = None
                for b in bancos_disponiveis.values():
                    nome_formatado = b['metodo'].upper() + " - " + b['nome']
                    if nome_formatado == item:
                        banco = b
                        break
                if banco is None:
                    QMessageBox.warning(None, "Erro", "Banco selecionado inválido.")
                    continue

                # Caixa para digitar o valor
                valor, ok = QInputDialog.getText(None, "Valor", "Digite o valor (ex: 1234 ou 1234,56):")
                if not ok or not valor:
                    QMessageBox.information(None, "Aviso", "Operação cancelada na entrada do valor.")
                    return
                if not re.fullmatch(r'\d+(,\d{1,2})?', valor):
                    QMessageBox.warning(None, "Erro", "Valor inválido.")
                    continue

                cnpj = banco["cnpj"]
                participante = banco["participante"]
                valor_formatado = valor.replace(",", ".")

                with open(self.arquivo_SPED, 'r+', encoding='latin1') as f:
                    linhas = f.readlines()

                registro_existente = next((l for l in linhas if l.startswith("|1601|") and l.split("|")[2] == cnpj), None)
                if registro_existente:
                    campos = registro_existente.strip().split("|")
                    valor_antigo_str = campos[4].replace(",", ".")
                    try:
                        valor_antigo = float(valor_antigo_str)
                        valor_novo = float(valor_formatado)
                        valor_total = valor_antigo + valor_novo
                        campos[4] = f"{valor_total:.2f}".replace(".", ",")
                        nova_linha = "|".join(campos) + "\n"
                        indice_linha = linhas.index(registro_existente)
                        linhas[indice_linha] = nova_linha
                        QMessageBox.information(None, "Sucesso", "Registro 1601 já existia — valor somado com sucesso!")
                    except ValueError:
                        QMessageBox.warning(None, "Erro", "Erro ao converter valores para soma.")
                else:
                    registro_0150 = (
                        f"|0150|{cnpj}|{participante['nome']}|{participante['cod_pais']}|{cnpj}|||{participante['cod_mun']}||"
                        f"{participante['logradouro']}|{participante['SN']}||{participante['bairro']}|\n"
                    )
                    registro_1601 = f"|1601|{cnpj}||{valor}|0|0|\n"
                    registro_1601 = self.garantir_campo_13_vazio(registro_1601)

                    if not any(f"|0150|{cnpj}|" in l for l in linhas):
                        pos_0100 = next((i for i, l in enumerate(linhas) if l.startswith("|0100|")), -1)
                        if pos_0100 != -1:
                            linhas.insert(pos_0100 + 1, registro_0150)
                            self.atualizar_bloco_9900(linhas, "0150", 1)
                            self.atualizar_bloco_0990(linhas, 1)

                    pos_1010 = next((i + 1 for i, l in enumerate(linhas) if l.startswith("|1010|")), -1)
                    if pos_1010 == -1:
                        QMessageBox.warning(None, "Erro", "Bloco |1010| não encontrado.")
                        return
                    
                    linhas = self.garantir_campo_13_vazio(linhas)
                    linhas.insert(pos_1010, registro_1601)
                    self.atualizar_bloco_9900(linhas, "1601", 1)
                    self.atualizar_bloco_1990(linhas, "1601")
                    self.substituir_bloco_1010(linhas)
                    self.atualizar_bloco_9999(linhas, sum(1 for l in linhas if l.strip()))
                    QMessageBox.information(None, "Sucesso", f"Registro 1601 para {banco['nome']} (CNPJ: {cnpj}) inserido com sucesso.")

                self.salvar_SPED(linhas)

                if cnpj not in cnpjs_ja_inseridos:
                    cnpjs_ja_inseridos.append(cnpj)

                continuar = QMessageBox.question(None, "Continuar?", "Deseja inserir outro banco para este cliente?", QMessageBox.Yes | QMessageBox.No)
                if continuar == QMessageBox.Yes:
                    continue
                else:
                    resposta_feito = QMessageBox.question(None, "Marcar como FEITO?", "Deseja marcar o cliente como FEITO?", QMessageBox.Yes | QMessageBox.No)
                    if resposta_feito == QMessageBox.Yes:
                        for cliente in self.clientes:
                            if cliente['codigo'] == cliente_selecionado['codigo']:
                                if self.edicao_liberada:
                                    cliente['status'] = 'FEITO'
                                    try:
                                        with open(self.caminho_clientes, 'w', newline='', encoding='latin1') as f:
                                            campos = self.clientes[0].keys()
                                            escritor = csv.DictWriter(f, fieldnames=campos)
                                            escritor.writeheader()
                                            escritor.writerows(self.clientes)
                                        QMessageBox.information(None, "Sucesso", "Status atualizado para FEITO com sucesso.")
                                    except Exception as e:
                                        QMessageBox.warning(None, "Erro", f"Erro ao atualizar status no arquivo clientes.csv: {e}")
                                else:
                                    try:
                                        self._registrar_pendencia_alterar_status(cliente['codigo'], "FEITO")
                                        QMessageBox.information(None, "Pendente", "Planilha bloqueada. Status salvo nas pendências.")
                                    except Exception as e:
                                        QMessageBox.warning(None, "Erro", f"Erro ao registrar pendência: {e}")
                                break

                    QMessageBox.information(None, "Finalizado", "Processo de inserção finalizado.")
                    break

        except Exception as e:
            QMessageBox.critical(None, "Erro", f"Ocorreu um erro inesperado: {e}")

    def inserir_registro_1601(self):
        # self.edicao_liberada = self._adquirir_lock()
        # if not self.edicao_liberada:
        #     QMessageBox.warning(None, "Somente leitura", "Outro usuário está editando a planilha. A inserção está bloqueada.")
        #     return
        try:
            # Seleciona o arquivo SPED dentro do diretório onde o código está
            arquivo_txt = self.selecionar_arquivo_SPED()
            if not arquivo_txt:
                input("\nOperação cancelada. Pressione Enter para continuar...")
                return

            self.arquivo_SPED = arquivo_txt

            print("=== Inserir Registro 1601 ===")
            
            # Carregar clientes
            caminho_clientes = os.path.join(DIRETORIO_SPED, "clientes.csv")
            try:
                with open(caminho_clientes, newline='', encoding='latin1') as f:
                    leitor = csv.DictReader(f)
                    self.clientes = list(leitor)
            except Exception as e:
                print(f"\nErro ao carregar clientes: {e}")
                input("Pressione Enter para continuar...")
                return
            
            if not self.clientes:
                print("\nNenhum cliente encontrado.")
                return

            # Carregar participantes
            caminho_participantes = os.path.join(DIRETORIO_SPED, "participantes.csv")
            try:
                with open(caminho_participantes, newline='', encoding='latin1') as f:
                    leitor = csv.DictReader(f)
                    self.participantes = list(leitor)
            except Exception as e:
                print(f"\nErro ao carregar participantes: {e}")
                input("Pressione Enter para continuar...")
                return


            print("\nClientes cadastrados:")
            for i, cliente in enumerate(self.clientes, 1):
                print(f"{i}. {cliente['nome_cliente']}")
                print(f"   Código: {cliente['codigo']}")
                print(f"   E-mail do contador: {cliente['email_contador']}")
                print(f"   Status: {cliente['status']}")
                print(f"   Meios de Pagamento: ")
                for metodo in ["pix_pdv", "pix_off", "pos_adiquirente", "boleto", "tef", "delivery"]:
                    banco_codigo = cliente.get(metodo)
                    if banco_codigo:
                        banco_info = next((p for p in self.participantes if p['cnpj'] == banco_codigo), None)
                        if banco_info:
                            print(f"     - {metodo.upper()}: Banco {banco_info.get('nome', 'N/I')} (CNPJ: {banco_codigo}) - Município: {banco_info.get('nome_mun', 'N/I')}")
                print('-' * 120)

            # --- Seleção do cliente via tela gráfica ---
            dialog = FormSelecionarCliente(self.clientes, self)
            if dialog.exec() == QDialog.Accepted:
                selected_item = dialog.list_widget.currentItem()
                if selected_item is None:
                    print("Nenhum cliente selecionado.")
                    return
                cliente_selecionado = selected_item.data(Qt.UserRole)
            else:
                print("Seleção cancelada.")
                return

            # --- Após selecionar o cliente, começa inserção dos 1601 ---
            cnpjs_ja_inseridos = []

            while True:
                meios_pagamento = ["pix_pdv", "pix_off", "pos_adiquirente", "boleto", "tef", "delivery"]
                bancos_disponiveis = {}
                indice_global = 1

                nao_registrados = []
                ja_registrados = []
                for metodo in meios_pagamento:
                    cnpj_banco = cliente_selecionado.get(metodo)
                    if cnpj_banco:
                        participante = next((p for p in self.participantes if p['cnpj'] == cnpj_banco), None)
                        if participante:
                            status = "JÁ INSERIDO" if cnpj_banco in cnpjs_ja_inseridos else "DISPONÍVEL"
                            entrada = {
                                "metodo": metodo,
                                "cnpj": cnpj_banco,
                                "nome": participante['nome'],
                                "municipio": participante['nome_mun'],
                                "participante": participante,
                                "status": status
                            }
                            if status == "DISPONÍVEL":
                                nao_registrados.append(entrada)
                            else:
                                ja_registrados.append(entrada)

                bancos_disponiveis = {}
                indice_global = 1

                lista_bancos = []

                if nao_registrados:
                    lista_bancos.append("--- Ainda não registrados ---")
                    for b in nao_registrados:
                        linha = f"{indice_global}. {b['metodo'].upper()}: {b['nome']} - CNPJ: {b['cnpj']} - Município: {b['municipio']} [DISPONÍVEL]"
                        lista_bancos.append(linha)
                        bancos_disponiveis[indice_global] = b
                        indice_global += 1

                if ja_registrados:
                    lista_bancos.append("--- Já inseridos anteriormente ---")
                    for b in ja_registrados:
                        linha = f"{indice_global}. {b['metodo'].upper()}: {b['nome']} - CNPJ: {b['cnpj']} - Município: {b['municipio']} [JÁ INSERIDO]"
                        lista_bancos.append(linha)
                        bancos_disponiveis[indice_global] = b
                        indice_global += 1

                # Caixa para seleção do banco
                itens_selecionaveis = [bancos_disponiveis[i]['metodo'].upper() + " - " + bancos_disponiveis[i]['nome'] for i in bancos_disponiveis]
                item, ok = QInputDialog.getItem(None, "Selecionar Banco", "Escolha o banco para inserir/somar valor:", itens_selecionaveis, 0, False)
                if not ok or not item:
                    QMessageBox.information(None, "Aviso", "Operação cancelada na seleção do banco.")
                    return

                # Encontrar o banco selecionado pelo nome+metodo
                banco = None
                for b in bancos_disponiveis.values():
                    nome_formatado = b['metodo'].upper() + " - " + b['nome']
                    if nome_formatado == item:
                        banco = b
                        break
                if banco is None:
                    QMessageBox.warning(None, "Erro", "Banco selecionado inválido.")
                    continue

                # Caixa para digitar o valor
                valor, ok = QInputDialog.getText(None, "Valor", "Digite o valor (ex: 1234 ou 1234,56):")
                if not ok or not valor:
                    QMessageBox.information(None, "Aviso", "Operação cancelada na entrada do valor.")
                    return
                if not re.fullmatch(r'\d+(,\d{1,2})?', valor):
                    QMessageBox.warning(None, "Erro", "Valor inválido.")
                    continue

                cnpj = banco["cnpj"]
                participante = banco["participante"]
                valor_formatado = valor.replace(",", ".")

                with open(self.arquivo_SPED, 'r+', encoding='latin1') as f:
                    linhas = f.readlines()

                registro_existente = next((l for l in linhas if l.startswith("|1601|") and l.split("|")[2] == cnpj), None)
                if registro_existente:
                    campos = registro_existente.strip().split("|")
                    valor_antigo_str = campos[4].replace(",", ".")
                    try:
                        valor_antigo = float(valor_antigo_str)
                        valor_novo = float(valor_formatado)
                        valor_total = valor_antigo + valor_novo
                        campos[4] = f"{valor_total:.2f}".replace(".", ",")
                        nova_linha = "|".join(campos) + "\n"
                        indice_linha = linhas.index(registro_existente)
                        linhas[indice_linha] = nova_linha
                        QMessageBox.information(None, "Sucesso", "Registro 1601 já existia — valor somado com sucesso!")
                    except ValueError:
                        QMessageBox.warning(None, "Erro", "Erro ao converter valores para soma.")
                else:
                    registro_0150 = (
                        f"|0150|{cnpj}|{participante['nome']}|{participante['cod_pais']}|{cnpj}|||{participante['cod_mun']}||"
                        f"{participante['logradouro']}|{participante['SN']}||{participante['bairro']}|\n"
                    )
                    registro_1601 = f"|1601|{cnpj}||{valor}|0|0|\n"
                    registro_1601 = self.garantir_campo_13_vazio(registro_1601)

                    if not any(f"|0150|{cnpj}|" in l for l in linhas):
                        pos_0100 = next((i for i, l in enumerate(linhas) if l.startswith("|0100|")), -1)
                        if pos_0100 != -1:
                            linhas.insert(pos_0100 + 1, registro_0150)
                            self.atualizar_bloco_9900(linhas, "0150", 1)
                            self.atualizar_bloco_0990(linhas, 1)

                    pos_1010 = next((i + 1 for i, l in enumerate(linhas) if l.startswith("|1010|")), -1)
                    if pos_1010 == -1:
                        QMessageBox.warning(None, "Erro", "Bloco |1010| não encontrado.")
                        return
                    
                    linhas = self.garantir_campo_13_vazio(linhas)
                    linhas.insert(pos_1010, registro_1601)
                    self.atualizar_bloco_9900(linhas, "1601", 1)
                    self.atualizar_bloco_1990(linhas, "1601")
                    self.substituir_bloco_1010(linhas)
                    self.atualizar_bloco_9999(linhas, sum(1 for l in linhas if l.strip()))
                    QMessageBox.information(None, "Sucesso", f"Registro 1601 para {banco['nome']} (CNPJ: {cnpj}) inserido com sucesso.")

                self.salvar_SPED(linhas)

                if cnpj not in cnpjs_ja_inseridos:
                    cnpjs_ja_inseridos.append(cnpj)

                continuar = QMessageBox.question(None, "Continuar?", "Deseja inserir outro banco para este cliente?", QMessageBox.Yes | QMessageBox.No)
                if continuar == QMessageBox.Yes:
                    continue
                else:
                    resposta_feito = QMessageBox.question(None, "Marcar como FEITO?", "Deseja marcar o cliente como FEITO?", QMessageBox.Yes | QMessageBox.No)
                    if resposta_feito == QMessageBox.Yes:
                        for cliente in self.clientes:
                            if cliente['codigo'] == cliente_selecionado['codigo']:
                                self.edicao_liberada = self._adquirir_lock()
                                if self.edicao_liberada:
                                    cliente['status'] = 'FEITO'
                                    try:
                                        with open(self.caminho_clientes, 'w', newline='', encoding='latin1') as f:
                                            campos = self.clientes[0].keys()
                                            escritor = csv.DictWriter(f, fieldnames=campos)
                                            escritor.writeheader()
                                            escritor.writerows(self.clientes)
                                        QMessageBox.information(None, "Sucesso", "Status atualizado para FEITO com sucesso.")
                                    except Exception as e:
                                        QMessageBox.warning(None, "Erro", f"Erro ao atualizar status no arquivo clientes.csv: {e}")
                                else:
                                    try:
                                        self._registrar_pendencia_alterar_status(cliente['codigo'], "FEITO")
                                        QMessageBox.information(None, "Pendente", "Planilha bloqueada. Status salvo nas pendências.")
                                    except Exception as e:
                                        QMessageBox.warning(None, "Erro", f"Erro ao registrar pendência: {e}")
                                break  # termina o for
                    QMessageBox.information(None, "Finalizado", "Processo de inserção finalizado.")
                    break
        except Exception as e:
            QMessageBox.critical(None, "Erro", f"Ocorreu um erro inesperado: {e}")

    def salvar_SPED(self, linhas):
        try:
            with open(self.arquivo_SPED, 'w', encoding='latin1') as f:
                f.writelines(linhas)
            print("✔ Arquivo SPED salvo com sucesso.")
        except Exception as e:
            print(f"❌ Erro ao salvar o SPED: {e}")

    def atualizar_status_cliente(self, cliente, novo_status):
        # Caminho SPED
        caminho_csv = os.path.join(DIRETORIO_SPED, 'clientes.csv')
        if not self.edicao_liberada:
            QMessageBox.warning(self, "Somente leitura", "Outro usuário está editando a planilha. Aguarde para atualizar o status.")
            return

        # Carrega novamente todos os clientes do arquivo
        try:
            with open(caminho_csv, newline='', encoding='latin1') as f:
                leitor = csv.DictReader(f)
                todos_clientes = list(leitor)
        except Exception as e:
            print(f"Erro ao recarregar clientes: {e}")
            return

        # Atualiza o status do cliente correto
        for c in todos_clientes:
            if c['codigo'] == cliente['codigo']:
                c['status'] = novo_status
                break

        # Reescreve o arquivo com os dados atualizados
        try:
            with open(caminho_csv, 'w', newline='', encoding='latin1') as f:
                campos = ['codigo', 'nome_cliente', 'email_contador', 'email_secundario', 'status',
                        'pix_pdv', 'pix_off', 'pos_adiquirente', 'boleto', 'tef', 'delivery']
                writer = csv.DictWriter(f, fieldnames=campos)
                writer.writeheader()
                writer.writerows(todos_clientes)
        except Exception as e:
            print(f"Erro ao salvar o arquivo de clientes: {e}")


    def atualizar_bloco_9900(self, linhas, registro, delta):
        for i, linha in enumerate(linhas):
            if linha.startswith(f"|9900|{registro}|"):
                partes = linha.strip().split("|")
                partes[3] = str(int(partes[3]) + delta)
                linhas[i] = "|".join(partes) + "\n"
                return

        # Se não existe ainda, inserir logo após o |9900|1010|
        pos_1010_9900 = next((i for i, l in enumerate(linhas) if l.startswith("|9900|1010|")), None)
        if pos_1010_9900 is not None:
            linhas.insert(pos_1010_9900 + 1, f"|9900|{registro}|{delta}|\n")
        else:
            # Se não encontrar |9900|1010|, insere no final do bloco 9900 como fallback
            pos_9900_fim = next((i for i, l in reversed(list(enumerate(linhas))) if l.startswith("|9900|")), len(linhas) - 1)
            linhas.insert(pos_9900_fim + 1, f"|9900|{registro}|{delta}|\n")

    def atualizar_bloco_0990(self, linhas, delta):
        qtd = sum(1 for l in linhas if l.startswith("|0"))
        for i, linha in enumerate(linhas):
            if linha.startswith("|0990|"):
                linhas[i] = f"|0990|{qtd}|\n"
                return

    def atualizar_bloco_1990(self, linhas, tipo):
        qtd = sum(1 for l in linhas if l.startswith("|1"))
        for i, linha in enumerate(linhas):
            if linha.startswith("|1990|"):
                linhas[i] = f"|1990|{qtd}|\n"
                return

    def atualizar_bloco_9999(self, linhas, total):
        total = sum(1 for l in linhas if l.strip())  # conta todas as linhas não vazias
        for i, linha in enumerate(linhas):
            if linha.startswith("|9999|"):
                linhas[i] = f"|9999|{total}|\n"
                return

    def substituir_bloco_1010(self, linhas):
        for i, linha in enumerate(linhas):
            if linha.startswith("|1010|N|N|N|N|N|"):
                # Substitui o bloco |1010|N|N|N|N|N| por |1010|N|N|N|N|N|N|S|N|N|N|N|N|N|
                linhas[i] = "|1010|N|N|N|N|N|N|S|N|N|N|N|N|N|\n"
        return linhas

    def garantir_campo_13_vazio(self, linhas):
        for i, linha in enumerate(linhas):
            if linha.startswith("|0200|"):
                campos = linha.split("|")
                if len(campos) > 13:
                    campos[13] = ""  # Deixa o campo 13 vazio
                    linhas[i] = "|".join(campos)  # Reconstrói a linha com o campo 13 vazio
        return linhas

    # Classe interna para o formulário de cadastro de participante
    class FormCadastrarParticipante(QDialog):
        def __init__(self, parent=None):
            super().__init__(parent)
            self.setWindowTitle("Cadastrar Participante")
            self.setModal(True)
            self.setFixedSize(600, 300)

            layout = QFormLayout(self)

            self.input_codigo = QLineEdit()
            self.input_codigo.setMaxLength(18)  # Aceita com máscara: 99.999.999/9999-99
            layout.addRow("*Código (CNPJ) [até 18 caracteres] :", self.input_codigo)

            self.btn_consultar = QPushButton("Consultar CNPJ")
            self.btn_consultar.clicked.connect(self.consultar_cnpj)
            layout.addRow("", self.btn_consultar)

            self.input_nome = QLineEdit()
            layout.addRow("*Nome do banco :", self.input_nome)

            self.input_cod_mun = QLineEdit()
            self.input_cod_mun.setMaxLength(10)
            layout.addRow("*Código do município (somente números) :", self.input_cod_mun)

            self.input_logradouro = QLineEdit()
            layout.addRow("*Logradouro :", self.input_logradouro)

            self.input_nome_mun = QLineEdit()
            layout.addRow("*Nome do município :", self.input_nome_mun)

            self.input_bairro = QLineEdit()
            layout.addRow("Bairro:", self.input_bairro)

            self.input_sn = QLineEdit()
            layout.addRow("Número:", self.input_sn)

            self.botoes = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            self.botoes.accepted.connect(self.validar_e_aceitar)
            self.botoes.rejected.connect(self.reject)
            layout.addWidget(self.botoes)

        def consultar_cnpj(self):
            cnpj = re.sub(r'\D', '', self.input_codigo.text().strip())  # Limpa máscara

            if len(cnpj) != 14:
                QMessageBox.warning(self, "Erro", "Digite um CNPJ válido (14 dígitos numéricos).")
                return

            try:
                url = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj}"
                response = requests.get(url)
                if response.status_code == 200:
                    dados = response.json()
                    self.input_nome.setText(dados.get("razao_social", ""))
                    self.input_logradouro.setText(dados.get("logradouro", ""))
                    self.input_nome_mun.setText(dados.get("municipio", ""))
                    self.input_cod_mun.setText(str(dados.get("codigo_municipio", "")))
                    self.input_bairro.setText(dados.get("bairro", ""))
                    self.input_sn.setText(dados.get("numero", ""))
                    QMessageBox.information(self, "Sucesso", "Dados preenchidos automaticamente.")
                else:
                    QMessageBox.warning(self, "Erro", f"CNPJ não encontrado. Código {response.status_code}")
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao consultar CNPJ: {str(e)}")

        def validar_e_aceitar(self):
            codigo = re.sub(r'\D', '', self.input_codigo.text().strip())  # Somente números do CNPJ
            cod_mun = re.sub(r'\D', '', self.input_cod_mun.text().strip())[:7]  # Somente números, máximo 7

            nome = self.input_nome.text().strip()
            logradouro = self.input_logradouro.text().strip()
            nome_mun = self.input_nome_mun.text().strip()
            bairro = self.input_bairro.text().strip() or "Centro"  # Se vazio, define como 'Centro'
            sn = self.input_sn.text().strip() or "0"  # Se vazio, define como '0'
            # endereco = self.input_endereco.text().strip()

            if len(codigo) != 14:
                QMessageBox.warning(self, "Erro", "O CNPJ deve conter exatamente 14 dígitos numéricos.")
                return

            if not nome:
                QMessageBox.warning(self, "Erro", "Nome do banco não pode ser vazio.")
                return

            if len(cod_mun) != 7:
                QMessageBox.warning(self, "Erro", "Código do município deve conter exatamente 7 dígitos numéricos.")
                return

            if not logradouro:
                QMessageBox.warning(self, "Erro", "Logradouro não pode ser vazio.")
                return

            if not nome_mun.replace(" ", "").isalpha():
                QMessageBox.warning(self, "Erro", "Nome do município deve conter apenas letras.")
                return

            if os.path.exists(ARQUIVO_PARTICIPANTES):
                with open(ARQUIVO_PARTICIPANTES, newline='', encoding="latin1") as f:
                    leitor = csv.DictReader(f)
                    for linha in leitor:
                        if linha.get("codigo", "") == codigo:
                            QMessageBox.warning(self, "Erro", f"Código {codigo} já está cadastrado.")
                            return

            self.novo_participante = {
                "codigo": codigo,
                "nome": nome,
                "cod_pais": "1058",
                "cnpj": codigo,
                "cod_mun": cod_mun,
                "logradouro": logradouro,
                "SN": sn,
                "bairro": bairro,
                # "endereco": endereco,
                "nome_mun": nome_mun
            }

            self.accept()

    # Classe interna para o formulário de cadastro de cliente
    class FormCadastrarCliente(QDialog):
        def __init__(self, parent=None):
            super().__init__(parent)
            self.setWindowTitle("Cadastrar Cliente")
            self.setModal(True)
            self.setFixedSize(500, 300)
            
            layout = QFormLayout(self)
            
            # Campos básicos
            self.input_codigo = QLineEdit()
            self.input_nome = QLineEdit()
            self.input_email_contador = QLineEdit()
            self.input_email_secundario = QLineEdit()
            
            layout.addRow("Código do Cliente (números):", self.input_codigo)
            layout.addRow("Nome do Cliente:", self.input_nome)
            layout.addRow("E-mail do Contador:", self.input_email_contador)
            layout.addRow("E-mail Secundário (opcional):", self.input_email_secundario)

            # Campos para meios de pagamento com ComboBox e filtro
            self.combo_pix_pdv = QComboBox()
            self.combo_pix_off = QComboBox()
            self.combo_pos_adiquirente = QComboBox()
            self.combo_boleto = QComboBox()
            self.combo_tef = QComboBox()
            self.combo_delivery = QComboBox()
            
            # Função local para popular o combo e configurar o filtro (QCompleter)
            def populate_combo(combo, participantes):
                combo.addItem("Nenhum", "")
                for p in participantes:
                    display = f"{p.get('nome', 'Sem Nome')} - {p.get('nome_mun', 'N/A')}"
                    combo.addItem(display, p.get("codigo", ""))
                # Torna o combo editável e configura o completer para filtragem
                combo.setEditable(True)
                combo.setInsertPolicy(QComboBox.NoInsert)
                items = [combo.itemText(i) for i in range(combo.count())]
                completer = QCompleter(items, combo)
                completer.setCaseSensitivity(Qt.CaseInsensitive)
                completer.setFilterMode(Qt.MatchContains)
                combo.setCompleter(completer)
            
            participantes = []
            if parent is not None and hasattr(parent, "participantes"):
                participantes = parent.participantes
            
            populate_combo(self.combo_pix_pdv, participantes)
            populate_combo(self.combo_pix_off, participantes)
            populate_combo(self.combo_pos_adiquirente, participantes)
            populate_combo(self.combo_boleto, participantes)
            populate_combo(self.combo_tef, participantes)
            populate_combo(self.combo_delivery, participantes)
            
            layout.addRow("Banco para PIX PDV:", self.combo_pix_pdv)
            layout.addRow("Banco para PIX OFF:", self.combo_pix_off)
            layout.addRow("Banco para POS ADIQUIRENTE:", self.combo_pos_adiquirente)
            layout.addRow("Banco para BOLETO:", self.combo_boleto)
            layout.addRow("Banco para TEF:", self.combo_tef)
            layout.addRow("Banco para DELIVERY:", self.combo_delivery)
            
            self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            self.buttons.accepted.connect(self.validar_e_aceitar)
            self.buttons.rejected.connect(self.reject)
            layout.addWidget(self.buttons)
        
        def validar_e_aceitar(self):
            codigo = self.input_codigo.text().strip()
            nome = self.input_nome.text().strip()
            email_contador = self.input_email_contador.text().strip()
            email_secundario = self.input_email_secundario.text().strip()
            
            if not codigo.isdigit():
                QMessageBox.warning(self, "Erro", "Código inválido, use apenas números.")
                return
            if len(nome) < 2 or not any(c.isalpha() for c in nome):
                QMessageBox.warning(self, "Erro", "Nome inválido, deve conter pelo menos duas letras.")
                return
            if "@" not in email_contador:
                QMessageBox.warning(self, "Erro", "E-mail do contador inválido, deve conter '@'.")
                return
            
            # Coleta as escolhas dos combos
            pix_pdv = self.combo_pix_pdv.currentData()
            pix_off = self.combo_pix_off.currentData()
            pos_adiquirente = self.combo_pos_adiquirente.currentData()
            boleto = self.combo_boleto.currentData()
            tef = self.combo_tef.currentData()
            delivery = self.combo_delivery.currentData()
            
            self.novo_cliente = {
                'codigo': codigo,
                'nome_cliente': nome,
                'email_contador': email_contador,
                'email_secundario': email_secundario,
                'status': "",
                'pix_pdv': pix_pdv,
                'pix_off': pix_off,
                'pos_adiquirente': pos_adiquirente,
                'boleto': boleto,
                'tef': tef,
                'delivery': delivery
            }
            self.accept()

class FormSelecionarCliente(QDialog):
    def __init__(self, clientes, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Selecionar Cliente")
        self.setFixedSize(400, 400)
        self.selected_client = None

        layout = QVBoxLayout(self)

        # Campo para filtrar os clientes
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Digite parte do nome para filtrar...")
        layout.addWidget(self.filter_edit)

        # Lista para exibir os clientes
        self.list_widget = QListWidget()
        self.list_widget.itemDoubleClicked.connect(self.on_double_click)
        layout.addWidget(self.list_widget)

        # Botões OK/Cancelar
        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        layout.addWidget(self.button_box)
        self.button_box.accepted.connect(self._on_accept)
        self.button_box.rejected.connect(self.reject)

        # Armazena a lista completa e popula a lista inicial
        self.all_clients = clientes
        self.populate_list(self.all_clients)
        self.filter_edit.textChanged.connect(self.filter_clients)

    def on_double_click(self, item):
        # Fecha o diálogo como se tivesse apertado OK
        self._on_accept()

    def populate_list(self, clients):
        self.list_widget.setUpdatesEnabled(False)
        self.list_widget.clear()
        for i, client in enumerate(clients):
            item_text = f"Código: {client['codigo']} - {client['nome_cliente']} - Status: {client.get('status', '')}"
            item = QListWidgetItem(item_text)
            item.setData(Qt.UserRole, client)
            self.list_widget.addItem(item)
            if i % 50 == 0:
                QApplication.processEvents()
        self.list_widget.setUpdatesEnabled(True)


    def filter_clients(self):
        """Filtra a lista de clientes pelo texto do filtro."""
        text = self.filter_edit.text().strip().lower()
        if not text:
            filtered = self.all_clients
        else:
            filtered = [client for client in self.all_clients if text in client['nome_cliente'].lower()]
        self.populate_list(filtered)

    def _on_accept(self):
        """Captura o cliente selecionado antes de fechar o diálogo."""
        item = self.list_widget.currentItem()
        if item:
            self.selected_client = item.data(Qt.UserRole)
            self.accept()
        else:
            QMessageBox.warning(self, "Seleção", "Selecione um cliente antes de confirmar.")

class FormSelecionarBanco(QDialog):
    def __init__(self, bancos, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Selecionar Banco")
        self.setFixedSize(400, 400)
        
        layout = QVBoxLayout(self)
        
        # Campo de filtro
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Digite parte do nome do banco para filtrar...")
        layout.addWidget(self.filter_edit)

        # Lista para exibir bancos disponíveis
        self.list_widget = QListWidget()
        self.list_widget.itemDoubleClicked.connect(self.on_double_click)
        layout.addWidget(self.list_widget)

        # Botões de confirmação e cancelamento
        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        layout.addWidget(self.button_box)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        # Armazena todos os bancos e preenche a lista inicial
        self.all_bancos = bancos
        self.populate_list(self.all_bancos)
        self.filter_edit.textChanged.connect(self.filter_bancos)

    def on_double_click(self, item):
        self.accept()

    def populate_list(self, bancos):
        """Popula a lista de bancos com base na lista filtrada"""
        self.list_widget.clear()
        for banco in bancos:
            item_text = f"{banco['nome']} - CNPJ: {banco['cnpj']} - Município: {banco['nome_mun']}"
            item = QListWidgetItem(item_text)
            item.setData(Qt.UserRole, banco)
            self.list_widget.addItem(item)

    def filter_bancos(self):
        """Filtra a lista de bancos conforme o usuário digita"""
        text = self.filter_edit.text().lower()
        filtered = [banco for banco in self.all_bancos if text in banco['nome'].lower()]
        self.populate_list(filtered)

class FormDetalhesParticipante(QDialog):
    def __init__(self, participante, parent=None, editar=False):
        super().__init__(parent)
        self.setWindowTitle("Detalhes do Participante")
        self.setFixedSize(400, 400)

        self.participante = participante
        self.editar = editar

        layout = QVBoxLayout(self)

        campos_legiveis = {
            "nome": "Nome",
            "cnpj": "CNPJ",
            "cod_pais": "Código do País",
            "cod_mun": "Código do Município",
            "nome_mun": "Nome do Município",
            "logradouro": "Logradouro",
            "bairro": "Bairro",
            "SN": "Número"
        }

        if self.editar:
            # modo edição: usa QLineEdit para cada campo
            form_layout = QFormLayout()
            layout.addLayout(form_layout)

            self.inputs = {}

            for chave, label_amigavel in campos_legiveis.items():
                valor = participante.get(chave, "")
                if chave == "cnpj":
                    valor = self.formatar_cnpj(valor)
                if chave == "cod_pais" and valor == "1058":
                    valor = "1058 - Brasil"

                line_edit = QLineEdit(valor)
                form_layout.addRow(label_amigavel + ":", line_edit)
                self.inputs[chave] = line_edit

        else:
            # modo visualização: só mostra labels com texto formatado
            for chave, label_amigavel in campos_legiveis.items():
                valor = participante.get(chave, "")
                if chave == "cnpj":
                    valor = self.formatar_cnpj(valor)
                if chave == "cod_pais" and valor == "1058":
                    valor = "1058 - Brasil"
                texto = f"<b>{label_amigavel}:</b> {valor}"
                label = QLabel(texto)
                label.setWordWrap(True)
                layout.addWidget(label)

        # Botões Ok / Cancelar só aparecem no modo edição
        if self.editar:
            button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            button_box.accepted.connect(self.accept)
            button_box.rejected.connect(self.reject)
        else:
            button_box = QDialogButtonBox(QDialogButtonBox.Ok)
            button_box.accepted.connect(self.accept)

        layout.addWidget(button_box)

    def formatar_cnpj(self, cnpj_raw):
        cnpj_num = ''.join(filter(str.isdigit, cnpj_raw))
        if len(cnpj_num) != 14:
            return cnpj_raw
        return f"{cnpj_num[:2]}.{cnpj_num[2:5]}.{cnpj_num[5:8]}/{cnpj_num[8:12]}-{cnpj_num[12:]}"
    
    def get_data(self):
        if self.editar:
            dados = {}
            for chave, line_edit in self.inputs.items():
                valor = line_edit.text().strip()
                # Se quiser pode adicionar formatação inversa no CNPJ aqui (remover máscara)
                dados[chave] = valor
            return dados
        else:
            return self.participante

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SPED1601GUI()
    window.show()
    sys.exit(app.exec())
