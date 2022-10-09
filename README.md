# Automatizacao de Processos
 
Projeto para aprender automatização de projetos usando python

## O que o projeto faz?

- Abre o opera
- Abre o drive (simulando um acesso a base de dados)
- Faz o download de uma planilha
- Calcula marcadores
- Abre o Gmail
- Envia um email com os marcadores e a planilha em anexo

## Como Usar

Obs: Como estou usando pyautogui para clicar em uma coordenada na tela, caso seu monitor seja de um tamanho diferente o programa não funcionará

Pré Requisitos:
- Estar logado no google com alguma conta qualquer (para o reconhecimento da pagina do drive carregada)
- Estar logado no Gmail com alguma conta qualquer (para enviar o email/reconhecimento da pagina carregada)

Após fazer o download do projeto, abra o arquivo 'automatizar.py'


```python
pyautogui.write("opera")     #16 nessa linha troque "opera" pelo seu navegador preferido
tabela = pd.read_excel(r"C:\Users\breno\Downloads\Vendas - Dez.xlsx")     #36 nessa linha troque o caminho para o caminho da sua pasta de Downloads
pyautogui.write("dummyemail.gmail.com")        #52 nessa linha troque por um email válido, caso queira receber o email
arquivo = r"C:\Users\breno\Downloads\Vendas - Dez.xlsx"      #71 nessa linha troque o caminho para o caminho da sua pasta de Downloads
```

Você também pode alterar as variavéis de "assunto" e "corpoDoEmail" caso queira trocar o conteúdo do email

basta abrir o seu terminal na pasta onde baixou o projeto e rodar
python automatizar.py
