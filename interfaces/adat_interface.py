import customtkinter as ctk
from tkinter import messagebox, filedialog
import threading


from PIL import Image
from customtkinter import CTkImage

import os
import shutil
import subprocess
import platform

from common.base_path import get_base_dir

from solutions.adat.attribute_full_A import attribute_full_A
from solutions.adat.attribute_full_AB import attribute_full_AB
from solutions.adat.attribute_mixed_AB import attribute_mixed_AB

class ADATFrame(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        BASE_DIR = get_base_dir()
        caminho_fundo = BASE_DIR / 'images' / 'fundo_data_analytics.png'

        # Fundo
        bg_image = Image.open(caminho_fundo).resize((640, 360), Image.Resampling.LANCZOS)
        bg_photo = CTkImage(light_image=bg_image, size=(640, 360))
        background_label = ctk.CTkLabel(self, image=bg_photo, text="")
        background_label.image = bg_photo  # manter referência
        background_label.place(relwidth=1, relheight=1)

        bg_color = "#000810"

        # Título
        ctk.CTkLabel(
            self,
            text="Data Analytics Solutions",
            font=("Calibri", 25, "bold"),
            text_color="white",
            bg_color=bg_color
        ).place(relx=0.5, y=35, anchor="n")

        # Subtítulo
        ctk.CTkLabel(
            self,
            text="Access Deprovisioning Automated Test",
            font=("Calibri", 18, "bold"),
            text_color="white",
            bg_color=bg_color
        ).place(relx=0.5, y=80, anchor="n")

        
        # Selecionar arquivo de input
        self.arquivo_input = None
        self.arquivo_output = None
        self.explorer = None

        ctk.CTkButton(
            self,
            text="Selecionar Arquivo",
            font=("Calibri", 12, "bold"),
            width=240,
            height=30,
            corner_radius=8,
            fg_color="#4CABF8",
            bg_color=bg_color,
            command=self.selecionar_arquivo
        ).place(relx=0.74, y=155, anchor="center")

        ctk.CTkButton(
            self,
            text="Limpar Arquivo",
            command=self.limpar_arquivo,
            width=240,
            height=20,
            font=("Calibri", 12, "bold"),
            fg_color="#4E4E4E",
            bg_color=bg_color,
            corner_radius=8,
            text_color="white"
        ).place(relx=0.74, y=185, anchor="center")
        

        self.label_arquivo = ctk.CTkLabel(self, text="Nenhum arquivo selecionado", font=("Calibri", 12), text_color="white", bg_color=bg_color)
        self.label_arquivo.place(relx=0.50, y=330, anchor="center")

        # Botões
        largura_botao = 270
        altura_botao = 50
        espacamento = 60
        y_base = 148  # mais centralizado
        font=("Calibri", 14, "bold")

        ctk.CTkButton(
            self,
            text="Testar Atributo A\n(Todos os Sistemas)",
            font=font,
            width=largura_botao,
            height=altura_botao,
            corner_radius=10,
            bg_color=bg_color,
            command=lambda: self.executar(1)
        ).place(relx=0.3, y=y_base, anchor="center")

        ctk.CTkButton(
            self,
            text="Testar Atributos A e B\n(Todos os Sistemas)",
            font=font,
            width=largura_botao,
            height=altura_botao,
            corner_radius=10,
            bg_color=bg_color,
            command=lambda: self.executar(2)
        ).place(relx=0.3, y=y_base + espacamento, anchor="center")

        ctk.CTkButton(
            self,
            text="Testar Atributos A e B\n(Conforme Informações Disponíveis)",
            font=font,
            width=largura_botao,
            height=altura_botao,
            corner_radius=10,
            bg_color=bg_color,
            command=lambda: self.executar(3)
        ).place(relx=0.3, y=y_base + 2 * espacamento, anchor="center")

        ctk.CTkButton(
            self,
            text="Soluções Disponíveis",
            font=("Calibri", 11, "bold"),
            width=120,
            height=25,
            corner_radius=10,
            fg_color="#4CABF8",
            hover_color="#4CABF8",
            bg_color=bg_color,
            text_color="white",
            command=self.voltar_home
        ).place(x=15, y=320)

        self.botao_download = ctk.CTkButton(
            self,
            text="Ver Resultado",
            font=("Calibri", 12, "bold"),
            width=200,
            height=28,
            corner_radius=8,
            fg_color="#27AE60",
            bg_color=bg_color,
            command=self.abrir_pasta_saida,
            state="disabled"
        )
        self.botao_download.place(relx=0.74, y=220, anchor="center")


    def executar(self, opcao):
        loading_popup = ctk.CTkToplevel(self)
        loading_popup.title("Status do Teste")
        loading_popup.geometry("300x100")
        loading_popup.resizable(False, False)
        loading_popup.transient(self.winfo_toplevel())
        loading_popup.lift()
        loading_popup.focus_force()
        loading_popup.configure(fg_color="#000810")
        largura_tela = loading_popup.winfo_screenwidth()
        altura_tela = loading_popup.winfo_screenheight()
        largura_popup = 300
        altura_popup = 100
        x = (largura_tela - largura_popup) // 2
        y = (altura_tela - altura_popup) // 2
        loading_popup.geometry(f"{largura_popup}x{altura_popup}+{x}+{y}")

        label_status = ctk.CTkLabel(loading_popup, text="⏳ Teste em andamento...", font=("Calibri", 14, "bold"))
        label_status.pack(expand=True, padx=20, pady=20)

        def run():
            try:
                if not self.arquivo_input:
                    loading_popup.after(0, lambda: label_status.configure(text="⚠️ Nenhum arquivo selecionado."))
                    return

                if opcao == 1:
                    self.explorer = attribute_full_A(self.arquivo_input)
                elif opcao == 2:
                    self.explorer = attribute_full_AB(self.arquivo_input)
                elif opcao == 3:
                    self.explorer = attribute_mixed_AB(self.arquivo_input)

                loading_popup.after(0, lambda: label_status.configure(text="✅ Teste Realizado!"))
                loading_popup.after(2000, loading_popup.destroy)

                self.botao_download.configure(state="normal")

            except Exception as e:
                loading_popup.after(0, lambda err=e: [
                    label_status.configure(text="❌ Erro ao executar."),
                    messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(err)}")
                ])

        threading.Thread(target=run, daemon=True).start()

    def voltar_home(self):
        from interfaces.home_interface import HomeFrame  # importação tardia para evitar circularidade
        self.controller.mostrar_frame(HomeFrame)


    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if caminho:
            self.arquivo_input = caminho
            nome = os.path.basename(caminho)
            self.label_arquivo.configure(text=f"{nome}")

    def limpar_arquivo(self):
        self.arquivo_input = None
        self.label_arquivo.configure(text="Nenhum arquivo selecionado")
    
    def abrir_pasta_saida(self):
        sistema = platform.system()
        if sistema == "Windows":
            subprocess.Popen(["explorer", str(self.explorer)])
        elif sistema == "Darwin":
            subprocess.Popen(["open", str(self.explorer)])
        else:
            subprocess.Popen(["xdg-open", str(self.explorer)])
