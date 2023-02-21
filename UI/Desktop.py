import customtkinter as ct
import yaml

ct.set_appearance_mode("System")
ct.set_default_color_theme("blue")


class App(ct.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.default_config: dict = yaml.safe_load(open("DesktopConfig.yaml", "r"))
        self.title(self.default_config['AppTitle'])
        self.geometry(f"{self.default_config['AppHeight']}x{self.default_config['AppWidth']}")
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure((0, 1, 2, 3), weight=0)
        self.grid_rowconfigure(4, weight=1)
        

if __name__ == "__main__":
    app = App()
    app.mainloop()
