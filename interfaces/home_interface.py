import customtkinter as ctk
from PIL import Image
from customtkinter import CTkImage

from common.base_path import get_base_dir

from interfaces.adat_interface import ADATFrame
from interfaces.dev_interface import DevFrame

class HomeFrame(ctk.CTkFrame):
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
            text="Soluções Disponíveis",
            font=("Calibri", 18, "bold"),
            text_color="white",
            bg_color=bg_color
        ).place(relx=0.5, y=80, anchor="n")

        # Botões (2x2)
        largura_botao = 250
        altura_botao = 38
        espacamento_vertical = 55
        relx = 0.50
        y_base = 148  # mais centralizado

        ctk.CTkButton(
            self,
            text="Access Deprovisioning Automated Test",
            font=("Calibri", 14, "bold"),
            width=largura_botao,
            height=altura_botao,
            corner_radius=10,
            bg_color=bg_color,
            command=lambda: controller.mostrar_frame(ADATFrame)
        ).place(relx=relx, y=y_base, anchor="center")

        ctk.CTkButton(
            self,
            text="Intelligent Sample Selection",
            font=("Calibri", 14, "bold"),
            width=largura_botao,
            height=altura_botao,
            corner_radius=10,
            bg_color=bg_color,
            command=lambda: controller.mostrar_frame(DevFrame)
        ).place(relx=relx, y=y_base + espacamento_vertical, anchor="center")

        ctk.CTkButton(
            self,
            text="SAR Review Population",
            font=("Calibri", 14, "bold"),
            width=largura_botao,
            height=altura_botao,
            corner_radius=10,
            bg_color=bg_color,
            command=lambda: controller.mostrar_frame(DevFrame)
        ).place(relx=relx, y=y_base + 2 * espacamento_vertical, anchor="center")

        ctk.CTkButton(
            self,
            text="LOG Review",
            font=("Calibri", 14, "bold"),
            width=largura_botao,
            height=altura_botao,
            corner_radius=10,
            bg_color=bg_color,
            command=lambda: controller.mostrar_frame(DevFrame)
        ).place(relx=relx, y=y_base + 3 * espacamento_vertical, anchor="center")
