from botcity.core import DesktopBot
import pandas as pd


# def abrir_contoso():

bot = DesktopBot()
path_app = "C:/Program Files (x86)/Contoso, Inc/Contoso Invoicing/LegacyInvoicingApp.exe"


bot.execute(path_app)
bot.wait(2000)
bot.maximize_window()

# Procurando pelo elemento 'invoices'
if not bot.find("invoices", matching=0.97, waiting_time=10000):
    not_found("invoices") # type: ignore
else:
    bot.click()


def not_found(label):
    print(f"Element not found: {label}")


def Cadastra_Faturas(data, conta, contato, valor, status):
    # Procurando pelo elemento 'novo_registro'
    if not bot.find("novo_registro", matching=0.97, waiting_time=10000):
        not_found("novo_registro")
    else:
        bot.click()

    # Procurando pelo elemento 'Date'
    if not bot.find("Date", matching=0.97, waiting_time=10000):
        not_found("Date")
    else:
        bot.click_relative(83, 7)
        bot.type_keys(['home'])
        bot.type_keys(['shift', 'end'])
        bot.paste(data)
        bot.tab()
        bot.paste(conta)
        bot.tab()
        bot.paste(contato)
        bot.tab()
        bot.paste(valor)

    # Procurando pelo elemento 'status'
    if not bot.find("status", matching=0.97, waiting_time=10000):
        not_found("status")
    else:
        bot.click_relative(79, 7)

    coluna = status

    # Verificando o valor da variável 'coluna' e clicando no item correto
    if coluna == 'Univoiced':
        if not bot.find("univoiced", matching=0.97, waiting_time=10000):
            not_found("univoiced")
        else:
            bot.click_relative(72, 25)

    elif coluna == 'Invoiced':
        if not bot.find("invoiced", matching=0.97, waiting_time=10000):
            not_found("invoiced")
        else:
            bot.click_relative(71, 49)

    else:
        if not bot.find("paid", matching=0.97, waiting_time=10000):
            not_found("paid")
        else:
            bot.click_relative(62, 70)


# Lendo os dados da planilha
dados = pd.read_excel(r"C:\Users\Ingrid\Treinamento-botcity\Projeto- BotCity\file\Contoso+Coffee+Shop+Invoices.xlsx")
print(dados)

# Criando o bot fora da função para ser reutilizado
bot = DesktopBot()

# Iterando sobre as linhas do DataFrame e chamando a função
for coluna in dados.itertuples(index=False):
    Cadastra_Faturas(str(coluna[0]), str(coluna[1]), str(coluna[2]), str(coluna[3]), str(coluna[4]))

# Após o loop, clicar no botão "salvar"
if not bot.find("salvar", matching=0.97, waiting_time=10000):
    print("Element not found: salvar")
else:
    bot.click()

# abrir_contoso()