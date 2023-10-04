function extrairTextoPDF() {
  var folderId = "Nome do arquivo"; // Substitua pelo ID da pasta
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByType("application/pdf");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("BD_DIFAL");

  sheet.getRange('A:D').clearContent(); // Limpa todo o conteúdo da aba BD_DIFAL
  
  var lastRow = 1; // Começa na linha 2
  
  while (files.hasNext()) {
    var file = files.next();
    var fileId = file.getId();
    
    var docFile = converterPDFparaDoc(fileId);
    
    if (docFile) {
      var docText = extrairTextoDoDoc(docFile.getId());
      
      if (docText) {
        var fileName = file.getName();
        var extractedText = extrairTextoDoNomeDoArquivo(fileName); // Extrair texto do nome do arquivo
        extractedText = extractedText.replace("NF", ""); // Remover "NF" da variável
        
        sheet.getRange(lastRow, 1).setValue(fileName);
        sheet.getRange(lastRow, 3).setValue(docText);
        sheet.getRange(lastRow, 2).setValue(extractedText); // Adicionar o texto extraído do nome do arquivo
        
        // Extrair os caracteres e realizar a conversão e divisão
        var extractedChars = docText.substring(8, 11) + docText.substring(12, 16);
        var extractedNumber = parseInt(extractedChars);
        var finalValue = extractedNumber / 100;
        
        sheet.getRange(lastRow, 4).setValue(finalValue); // Adicionar o valor calculado
        
        /*Testando nova coluna 
        var value = lastRow.toString();
        var formattedValue = Utilities.formatString("%03d", value);
        sheet.getRange(lastRow, 5).setValue(formattedValue);
        */
        lastRow++;
        
        Logger.log("Texto extraído do documento '" + fileName + "': " + docText);
        
        // Excluir o arquivo .gdoc convertido
        DriveApp.getFileById(docFile.getId()).setTrashed(true);
      } else {
        Logger.log("Falha ao extrair texto do documento");
      }
    } else {
      Logger.log("Falha ao converter PDF para documento");
    }
  }
}

function extrairTextoDoNomeDoArquivo(fileName) {
  var delimiters = /-/; // Use um delimitador que separa os elementos
  var parts = fileName.split(delimiters);

  if (parts.length >= 3) {
    return parts[1]; // O texto entre os delimitadores é o segundo elemento
  } else {
    return "";
  }
}

function converterPDFparaDoc(fileId) {
  var url = "https://www.googleapis.com/drive/v3/files/" + fileId + "/copy";
  
  var payload = {
import pandas as pd
import pyautogui
import time
import keyboard
from datetime import datetime
import os
import threading  # Adicionando import para threading

caminho_do_arquivo = os.path.realpath(__file__)
caminho_do_arquivo = os.path.dirname(caminho_do_arquivo)
nome_arquivo = caminho_do_arquivo + "/" + "Titulos.xlsx"

# Leitura do arquivo Excel
Df = pd.read_excel(
    nome_arquivo, engine="openpyxl", dtype={'NATUREZA': str, 'FORNECEDOR': str, 'CODIGO_BARRAS': str, 'NO. TITULO': str})

# Variável de controle para interromper o script
script_executando = True

# Intervalo de espera entre as digitações (em segundos)
intervalo_digitar = 0.02  # Altere este valor conforme necessário

# Intervalo de espera entre as ativações de hotkey (em segundos)
intervalo_hotkey = time.sleep(0.05)

# Função para formatar o valor como string com duas casas decimais e zeros à direita


def formatar_valor(valor):
    valor_formatado = f'{valor:.2f}'
    if '.' in valor_formatado:
        parte_decimal = valor_formatado.split('.')[1]
        if len(parte_decimal) == 1:
            valor_formatado += '0'
    else:
        valor_formatado += '.00'
    return valor_formatado

# Função para executar a automação em cada linha do DataFrame


def realizar_automacao(row):
    global script_executando

    # Verifica se o script deve ser interrompido
    if not script_executando:
        return

    # Espera 1 segundo antes de iniciar a próxima ação
    time.sleep(0.5)

    # Simula a digitação do Prefixo
    pyautogui.write(str(row['PREFIXO']), interval=intervalo_digitar)
    intervalo_hotkey

    # Move para a próxima coluna usando a tecla "Tab"
    pyautogui.press('tab')
    intervalo_hotkey

    
    # Simula a digitação do NumTitulo
    pyautogui.write(str(row['NO. TITULO']), interval=intervalo_digitar)
    intervalo_hotkey

     # Move para a próxima coluna usando a tecla "Tab"
    pyautogui.press('tab')
    intervalo_hotkey
    
    # Simula a digitação do Tipo
    pyautogui.write(str(row['TIPO']), interval=intervalo_digitar)
    intervalo_hotkey
    
    # Simula a digitação da Natureza
    pyautogui.write(str(row['NATUREZA']), interval=intervalo_digitar)
    intervalo_hotkey
    
    # Move para a próxima coluna usando a tecla "Tab"
    pyautogui.press('tab')
    intervalo_hotkey
    
    # Simula a digitação do Fornecedor
    pyautogui.write(str(row['FORNECEDOR']), interval=intervalo_digitar)
    intervalo_hotkey
    
    # Move para a próxima coluna usando a tecla "Tab"O
    
    pyautogui.press('tab')  # Duas vezes para chegar à próxima coluna
    intervalo_hotkey

    pyautogui.press('tab')  # Duas vezes para chegar à próxima coluna
    intervalo_hotkey   

    # Formata a data em DD/MM/AAAA
    vencto_real = datetime.strptime(str(row['VENCTO_REAL']), '%Y-%m-%d %H:%M:%S').strftime('%d/%m/%Y')
    
    # Simula a digitação do Vencto_Real formatado
    pyautogui.write(vencto_real, interval=intervalo_digitar)
    intervalo_hotkey
    
    # Move para a próxima coluna usando a tecla "Tab"
    pyautogui.press('tab')
    intervalo_hotkey
    
    # Formata o valor de VLR_TITULO como string com duas casas decimais e zeros à direita
    valor_formatado = formatar_valor(row['VLR_TITULO'])
    
    # Simula a digitação do Vlr_Titulo formatado
    pyautogui.write(valor_formatado, interval=intervalo_digitar)
    intervalo_hotkey
    
    # Simula a digitação do Historico
    pyautogui.write(str(row['HISTORICO']), interval=intervalo_digitar)
    intervalo_hotkey    

    # Simula o atalho para salvar (Ctrl + S)
    pyautogui.hotkey('alt', 'b')
    time.sleep(1.5)

    # Move para a próxima coluna usando a tecla "Tab"
    pyautogui.press('tab')
    intervalo_hotkey

    # Move para a próxima coluna usando a tecla "Tab"
    pyautogui.press('tab')
    intervalo_hotkey

    # Move para a próxima coluna usando a tecla "Tab"
    pyautogui.press('tab')
    intervalo_hotkey
    
    # Simula a digitação do Codigo de barras
    pyautogui.write(str(row['CODIGO_BARRAS']), interval=intervalo_digitar)
    intervalo_hotkey

    pyautogui.hotkey('alt', 'u')
    time.sleep(1.5)

    pyautogui.write(str('22'), interval=intervalo_digitar)     
    intervalo_hotkey

    # Simula o atalho para salvar (Ctrl + S)
    pyautogui.hotkey('ctrl', 's')

    # Espera 6 segundos antes de continuar com a próxima linha
    time.sleep(4)

# Função para interromper o script quando a tecla "Esc" for pressionada


def interromper_script(e):
    global script_executando
    script_executando = False
    print("Script interrompido")

# Função para executar a automação em loop enquanto o script estiver ativo


def automacao_loop():
    for _, row in Df.iterrows():
        if not script_executando:
            break
        realizar_automacao(row)


# Define a combinação de teclas para ativar a automação
teclas_ativacao = "ctrl+alt+shift"

# Registra um manipulador de eventos para a combinação de teclas
keyboard.add_hotkey(teclas_ativacao, lambda: threading.Thread(
    target=automacao_loop).start())

# Registra um manipulador de eventos para a tecla "Esc" para interromper o script
keyboard.on_press_key('esc', interromper_script)    

try:
    # Mantém o script em execução
    keyboard.wait()
except KeyboardInterrupt:
    keyboard.on_press_key('esc', interromper_script)
    pass
finally:
    # Limpa todos os manipuladores de eventos
    keyboard.unhook_all()
