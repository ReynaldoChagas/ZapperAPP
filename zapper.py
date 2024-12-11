from tkinter import *
from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import WebDriverException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
from message import Message


class Window(Tk):
    def __init__(self):
        super().__init__()
        self.df = None
        self.excel_path = StringVar(self)
        self.model = StringVar(self)
        self.status = StringVar(self)
        self.column_var = StringVar(self)
        self.geometry('680x550')
        self.title('Zapper')
        self.resizable(FALSE,FALSE)
        self.driver = None


    def verifica_status(self,label_img,label_status,check,uncheck):
        try:
            self.driver.find_element(By.XPATH,'//*[@id="side"]/div[1]/div/div[2]/div[2]/div/div[1]')
            label_img.config(image = check)
            label_status.config(text='Connected')
        except WebDriverException:
            label_img.config(image = uncheck)
            label_status.config(text='Disconnected')
        except AttributeError:
            label_img.config(image = uncheck)
            label_status.config(text='Disconnected')
        finally:
            self.after(3000, lambda: self.verifica_status(label_img,label_status,check,uncheck))

    def choose_excel(self,menu):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", ".xlsx;.xls")]
        )
        if not file_path:
            return
        # Lê a planilha usando pandas
        self.df = pd.read_excel(file_path)
        if 'Status' not in self.df.columns:
            self.df['Status'] = ''
        self.df = self.df.fillna('').astype(str)
        self.excel_path.set(file_path.split('/')[-1])
        # Atualiza o menu suspenso com os nomes das colunas
        column_names = self.df.columns.tolist()
        self.column_var.set(column_names[0])  # Define o valor padrão
        menu['values'] = column_names     # Atualiza os valores do menu suspenso

    def choose_model(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Text files", ".txt")]
        )
        if not file_path:
            return
        #Lê o modelo da mensagem no formato txt
        with open(file_path,'r',encoding='utf-8') as file:
            self.model.set(file.read())

  
    def login_whatsapp(self):
        self.driver = webdriver.Chrome()
        self.driver.get('https://web.whatsapp.com/')
        print('Verificando login')
        while True:
            try:
                self.driver.find_element(By.XPATH,'//*[@id="side"]/div[1]/div/div[2]/div[2]/div/div[1]')
                print("Whatsapp logado.")
                break #whatsapp já está logado
            except NoSuchElementException:
                sleep(5) #tempo para aguardar o login
        sleep(3) #tempo para carregar a página após o login

    def save_excel(self):
        self.df.to_excel('results.xlsx',index=False)

    def start_program(self):
        try:
            message = Message(self.model.get())
            for row in self.df.index:
                message.update_text(self.df,row)
                contato = self.df.loc[row,self.column_var.get()] #define o contato para enviar a mensagem
                if contato == '':
                    self.df.loc[row,'Status'] = 'Sem número para contato'
                    continue #evita loop infinito qundo o contato é vazio
                else:
                    try:
                        self.driver.get(f'https://web.whatsapp.com/send/?phone=55${contato}&text={message.url_text}')
                        campo_texto = WebDriverWait(self.driver,15).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/div/div[1]/p')))
                        sleep(2)
                        campo_texto.send_keys(Keys.ENTER) 
                        self.df.loc[row,'Status'] = 'Enviado'
                        sleep(5)
                    except TimeoutException: #Exceção para erros de timeout
                        self.df.loc[row,'Status'] = 'Verificar contato.'
                        continue
                    except StaleElementReferenceException: #Exceção para erros de timeout
                        self.df.loc[row,'Status'] = 'Falha no envio.'
                        continue
                    except Exception as e:
                        print(f'Erro {e} ao enviar para {contato}.\nAplicação encerrada. Entre em contato com o administrador') #detecta outros possíveis erros para posterior tratamento
                        break 
        except Exception as e:
            print(f'Erro. {e}, entre em contato com o administrador')
        finally:
            self.save_excel()



window = Window()
icon = PhotoImage(file='resources/zapper.png')
window.iconphoto(False, icon)
window.update_idletasks()

bg = 'lightblue'
pady = 10
padx = 10
font = ('Comic Sans MS',10,'bold')

#----------------------------------
# Frame Principal

frame_principal = Frame(window,bd=2,relief='raised')

check = PhotoImage(file='resources/check.png')
uncheck = PhotoImage(file='resources/uncheck.png')
label_imagem = Label(frame_principal)
label_status = Label(frame_principal,font= font,text= 'Disconnected')
zapper_label = Label(frame_principal,font= font,text= 'Zapper\nBy: ReynaldoChagas',height= 14,anchor= 's')
window.verifica_status(label_imagem,label_status,check,uncheck)

button_excel = Button(frame_principal,font= font,text = 'Selecione um arquivo excel:',command= lambda: window.choose_excel(column_menu))
button_model = Button(frame_principal,font= font,text = 'Selecione um arquivo de texto:',command= window.choose_model)
button_login = Button(frame_principal,font= font,text = 'Login Whatsapp',command= window.login_whatsapp)
button_start = Button(frame_principal,font= font,text = 'Disparar mensagens',command= window.start_program)

button_excel.grid(row=0,column=0,pady=pady,padx=padx,sticky='we')
button_model.grid(row=1,column=0,pady=pady,padx=padx,sticky='we')
button_login.grid(row=2,column=0,pady=pady,padx=padx,sticky='we')
label_status.grid(row=3,column=0,padx=padx)
label_imagem.grid(row=4,column=0,padx=padx)
button_start.grid(row=5,column=0,pady=pady,padx=padx,sticky='we')
zapper_label.grid(row=6,column=0,pady=pady,padx=padx,sticky='we')

#----------------------------------
# Frame Secundario (EXCEL E MODEL)

#Excel
frame_secundario = Frame(window,bg=bg)
label1 = Label(frame_secundario,font= font,text = 'Planilha selecionada:',bg=bg,height=2,anchor='s')
label_excel = Label(frame_secundario,font= font, textvariable = window.excel_path,bg=bg,width=55,height=1,foreground='blue')
label2 = Label(frame_secundario,font= font,text = 'Selecione a coluna com os contatos:',bg=bg,height=2,anchor='s')
column_menu = ttk.Combobox(frame_secundario,font= font, textvariable=window.column_var,foreground='blue')

#Model
label3 = Label(frame_secundario,font= font,text= 'Modelo de mensagem:',bg='lightblue')
label_model = Label(frame_secundario,font= font,textvariable = window.model,bg=bg,justify='left',wraplength=440,foreground='blue')

#Grid
label1.grid(row=0)
label_excel.grid(row=1)
label2.grid(row=2)
column_menu.grid(row=3)
label3.grid(row=4,pady=pady)
label_model.grid(row=5)

#----------------------------------
# Grid Window

frame_principal.grid(row=0,column=0,padx=1)
frame_secundario.grid(row=0,column=1,sticky='nsew')
window.mainloop()

