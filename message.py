import urllib

class Message:
    def __init__(self,model):
        self.__model = str(model)
        self.__text = None
        self.__url_text = None
        
    @property
    def model(self):
        return self.__model
    
    @property
    def url_text(self):
        return self.__url_text

    @property
    def text(self):
         return self.__text
    
    @model.setter
    def model(self,model):
        self.__model = str(model)

    def _update_url(self):
        self.__url_text = urllib.parse.quote(self.__text)
  
    def update_text(self,df,row): #Atualiza o texto para cada linha da base de dados
        itens = list(map(lambda x:'{' + x + '}',df.columns))
        self.__text = self.__model
        for i,item in enumerate(itens):
            self.__text = self.__text.replace(item,df.loc[row,df.columns[i]])
        self._update_url() #Atualiza o atributo url da mensagem