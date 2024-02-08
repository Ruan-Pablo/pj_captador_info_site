from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl

site = "https://m.magazineluiza.com.br/busca/notebook/"
 
# Acessar o site:
driver = webdriver.Edge()
driver.get(site)
# extrair todos os títulos
#   Usando o selenium, encontra atraves do inspecionar - com o EXPATH//tag[@attribut=valor] no caso de classes com mais de um atributo não sei como faz

titulos = driver.find_elements(By.XPATH, "//h2[@data-testid='product-title']")
print(titulos)
precos = driver.find_elements(By.XPATH, "//p[@data-testid='price-value']")
print(precos)

# Criando a planilha
workbook = openpyxl.Workbook()
# Criando a página 'produtos'
workbook.create_sheet('produtos-note')
# Seleciono a página produtos
sheet_produtos = workbook['produtos-note']
sheet_produtos['A1'].value = 'Produto'
sheet_produtos['B1'].value = 'Preço'


# inserir os títulos e preços na planilha
for titulo, preco in zip(titulos, precos):
    sheet_produtos.append([titulo.text,preco.text])

workbook.save('produtos.xlsx')
