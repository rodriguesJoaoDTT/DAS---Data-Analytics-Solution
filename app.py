import customtkinter as ctk

from interfaces.home_interface import HomeFrame
from interfaces.adat_interface import ADATFrame
from interfaces.dev_interface import DevFrame

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Data Analytics Solutions")
        self.geometry("640x360")
        self.resizable(False, False)

        # Centralizar na tela
        self.update_idletasks()
        largura_tela = self.winfo_screenwidth()
        altura_tela = self.winfo_screenheight()
        x = (largura_tela - 640) // 2
        y = (altura_tela - 360) // 2
        self.geometry(f"640x360+{x}+{y}")

        # Configurar grid da janela principal
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Container para os frames
        self.container = ctk.CTkFrame(self, fg_color="transparent")
        self.container.grid(row=0, column=0, sticky="nsew")

        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        self.frames = {}

        # Criar todos os frames uma vez
        for FrameClass in (HomeFrame, ADATFrame, DevFrame):
            frame = FrameClass(self.container, self)
            self.frames[FrameClass] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.mostrar_frame(HomeFrame)

    def mostrar_frame(self, FrameClass):
        frame = self.frames[FrameClass]
        frame.tkraise()
