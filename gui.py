import threading
from pathlib import Path
from tkinter import Tk, Canvas, Entry, Text, Button, PhotoImage
from functionsB import *
import tkinter as tk
from tkinter import PhotoImage, scrolledtext
import sys
from functionsB import TextRedirector
from functionsB import run_automation
from functionsB import parar
import pygame




OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / "assets" / "frame0"
# def refreshsheet(): 
    # print("refresh")
#
def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)



pygame.mixer.init()
def load_sound(file_path):
    return pygame.mixer.Sound(file_path)

# Função para tocar o áudio
def play_sound(sound):
    sound.play()
    
sound = load_sound(relative_to_assets("amongus.mp3"))   
    
def criar_janela():
    global entry_2
    root = Tk("Automação Mensagens")
    root.title("Automação CHATBOT")
    root.geometry("600x554")
    root.configure(bg = "#212229")
    
    
    def on_button_click():
        play_sound(sound)
    
    
    
    def start_automation():
        # Função para iniciar a automação em uma thread separada
        def run_automation_thread():
            limite_mensagens = int(entry_2.get())
            try:
                run_automation(limite_mensagens)
            except Exception as e:
                print(f"Erro durante a automação: {e}")

        # Criar e iniciar a thread
        threading.Thread(target=run_automation_thread).start()

    #Janela principal
    photo = tk.PhotoImage(file=relative_to_assets("automacao.png"))
    root.wm_iconphoto(False, photo)
    canvas = Canvas(
        root,
        bg = "#212229",
        height = 554,
        width = 600,
        bd = 0,
        highlightthickness = 0,
        relief = "ridge"
        
    )
    canvas.place(x = 0, y = 0)

    #TextBox LOG
    entry_image_1 = PhotoImage(
        file=relative_to_assets("entry_1.png"))

    entry_bg_1 = canvas.create_image(
        299,
        195,
        image=entry_image_1
    )

    entry_1 = scrolledtext.ScrolledText(
        bd= 3,
        bg="#A8A8A8",
        fg="#000716",
        highlightthickness=0
    )
    entry_1.place(
        x=10.0,
        y=13.0,
        width=580,
        height=365
    )
    entry_1.config(state='disabled')
    
    text_redirector = TextRedirector(entry_1, sys.stdout)
    sys.stdout = text_redirector

        
    #Botão INICIAR
    button_image_1 = PhotoImage(
        file=relative_to_assets("button_1.png"))
    
    button_1 = Button(
        image=button_image_1,
        borderwidth=0,
        highlightthickness=0,
        command= start_automation,
        relief="flat"
    )

    button_1.place(
        x=450.0,
        y=483.0,
        width=125.0,
        height=45
    )
    #entry Limite
    
    
    entry_image_2 = PhotoImage(
        file=relative_to_assets("entry_2.png"))

    entry_bg_2 = canvas.create_image(
        543.5,
        451.0,
        image=entry_image_2
    )
    entry_2 = Entry(
        bd= 0,
        bg="#D9D9D9",
        fg="#000716",
        highlightthickness=0
    )
    entry_2.place(
        x=518.0,
        y=435.0,
        width=55.0,
        height=30.0
    )
    canvas.create_text(
        417.0,
        415.0,
        anchor="nw",
        text="Limite de mensagens:\n",
        fill="#FFFFFF",
        font=("Inter", 15 * -1)
    )

    openSheet = canvas.create_text(
        28.0,
        483.0,
        anchor="nw",
        text="ABRIR A PLANILHA GOOGLE",
        fill="#4a90e2",
        font=("Inter", 15 * -1),
        tags="sheet_link"
    )

    howToUse = canvas.create_text(
        
        28.0,
        502.0,
        anchor="nw",
        text="COMO USAR",
        fill="#FFFFFF",
        font=("Inter", 15 * -1),
        tags="howtouse_text"
    )
    image_image_1 = PhotoImage(
    file=relative_to_assets("image_1.png"))

    amongus = Button(
        image=image_image_1,
        borderwidth=0,
        highlightthickness=0,
        command= on_button_click,
        relief="flat",
        bg= "#212229"
    )

    amongus.place(
        x=10.0,
        y=379.0,
        width=18,
        height=18
    )
    
    
    # #AMONGUS
    # image_image_1 = PhotoImage(
    #     file=relative_to_assets("image_1.png"))

    # image_1 = canvas.create_image(
    #     21.0,
    #     389.0,
    #     image=image_image_1
    # )
    
    
    # refresh_image = PhotoImage(
    #     file=relative_to_assets("refresh_1.png"))
    
    # refresh_button_1 = Button(
    #     image=refresh_image,
    #     borderwidth=0,
    #     highlightthickness=0,
    #     command= refreshsheet,
    #     relief="flat",
    #     bg= "#212229"
    # )

    # refresh_button_1.place(
    #     x=470.0,
    #     y=435.0,
    #     width=32,
    #     height=32
    # )
    
   
    
    #Botão PARAR
    button_image_2 = PhotoImage(
        file=relative_to_assets("button_2.png"))

    button_2 = Button(
        image=button_image_2,
        borderwidth=0,
        highlightthickness=0,
        command=parar,
        relief="flat"
    )

    button_2.place(
        x=292.0,
        y=483.0,
        width=125.0,
        height=43.0
    )
    def on_canvas_click(event):
        # Verificar se o clique está na área do texto
        bbox_open_sheet = canvas.bbox("sheet_link")
        bbox_how_to_use = canvas.bbox("howtouse_text")
        
        if bbox_open_sheet and bbox_open_sheet[0] <= event.x <= bbox_open_sheet[2] and bbox_open_sheet[1] <= event.y <= bbox_open_sheet[3]:
            open_google_sheet()
        elif bbox_how_to_use and bbox_how_to_use[0] <= event.x <= bbox_how_to_use[2] and bbox_how_to_use[1] <= event.y <= bbox_how_to_use[3]:
            show_how_to_use()
            
    def on_canvas_enter(event):
        bbox_open_sheet = canvas.bbox("sheet_link")
        bbox_how_to_use = canvas.bbox("howtouse_text")

        if bbox_open_sheet and bbox_open_sheet[0] <= event.x <= bbox_open_sheet[2] and bbox_open_sheet[1] <= event.y <= bbox_open_sheet[3]:
            canvas.config(cursor="hand2")
        elif bbox_how_to_use and bbox_how_to_use[0] <= event.x <= bbox_how_to_use[2] and bbox_how_to_use[1] <= event.y <= bbox_how_to_use[3]:
            canvas.config(cursor="hand2")
        else:
            canvas.config(cursor="arrow")

    def on_canvas_leave(event):
        canvas.config(cursor="arrow")
        
                  

    # Associar o evento de clique ao Canvas
    canvas.bind("<Button-1>", on_canvas_click)
    canvas.bind("<Motion>", on_canvas_enter)
    canvas.bind("<Leave>", on_canvas_leave)
          
    root.resizable(False, False)
    root.mainloop()
    
    
    
    
#executar se abrir pelo Gui.py    
criar_janela()