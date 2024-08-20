# Defina o dicionário
dicionario = {
    "DISPONIVEL": "Laudo técnico: Equipamento disponível conforme checklist padrão Limpeza: Realizado a limpeza externa da carcaça, teclado e Tela",
    "DESCARTE": "De acordo com o padrão estabelecido pelo Itech e encerramento da garantia, o equipamento apresenta defeito / avaria no (Placa Mae), sendo inviavel o retorno do parque. Em função do equipamento estar obsoleto e não haver possibilidade de troca da peça, sugerimos o descarte. Foram removidas as peças Tela, Teclado, Bateria, Antena e Placa Wi-Fi, bateria da Bios, Sperker,  Cooler de Refrigeçao,   Memoria Ram, TouchPad, SSD:",
    "BATERIA": "De acordo com o padrão estabelecido pelo Itech, a bateria apresentou desempenho de (). , sendo necessária a troca da mesma.",
    "DEFEITO": "De acordo com o padrão estabelecido pelo Itech, as peças ( peças com defeito aqui ) apresentaram defeito, sendo necessária a troca da mesma.",
    "FABRICANTE": "Realizado teste no equipamento, foi percebido defeito ( ) Feito pedido de peças junto ao fabricante Ticket de reparo ( )",
    "DOACAO": "De acordo com o padrão estabelecido pelo Itech e encerramento da garantia, o equipamento apresenta baixo desempenho durante os testes de hardware, sendo inviavel o retorno do parque. Em função do equipamento estar obsoleto e não haver possibilidade de troca da peça, sugerimos o descarte."
}

# Crie o arquivo TXT
with open("Res_cliente.txt", "w") as arquivo:
    for chave, valor in dicionario.items():
        arquivo.write(f"{chave}: {valor}\n")

print("Dicionário salvo em dicionario.txt")