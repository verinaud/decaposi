import tkinter as tk
from pywinauto import Application

class Mensagem:
    def __init__(self, mensagem):
        self.root = tk.Tk()
        self.root.title("Aviso")

        self.label = tk.Label(self.root, text=mensagem, padx=10, pady=10)
        self.label.pack()

        self.botao_ok = tk.Button(self.root, text="OK", command=self.fechar_tela)
        self.botao_ok.pack(pady=10)

        self.root.bind("<Return>", lambda event: self.fechar_tela())

        self.root.mainloop()

        try: 
            app = Application().connect(title_re="^Aviso.*")
            dlg = app.window(title_re="^Aviso.*")
        except Exception:
            pass
        
        
    def fechar_tela(self):
        self.root.destroy()
        return True
    

