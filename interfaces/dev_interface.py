import customtkinter as ctk
from PIL import Image
from customtkinter import CTkImage

from common.base_path import get_base_dir

class DevFrame(ctk.CTkFrame):
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
            text="⏳ Solução em Desenvolvimento",
            font=("Calibri", 26, "bold"),
            text_color="white",
            bg_color=bg_color
        ).place(relx=0.5, y=167, anchor="n")

        ctk.CTkButton(
            self,
            text="Soluções Disponíveis",
            font=("Calibri", 11, "bold"),
            width=160,
            height=25,
            corner_radius=10,
            fg_color="#4CABF8",
            hover_color="#4CABF8",
            bg_color=bg_color,
            text_color="white",
            command=self.voltar_home
        ).place(x=20, y=315)

    def voltar_home(self):
        from interfaces.home_interface import HomeFrame
        self.controller.mostrar_frame(HomeFrame)
