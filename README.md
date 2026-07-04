# tempo.exe
Aplicativo de previsão do tempo que roda em ambiente Ms-DOS

O aplicativo foi gerado em Visual Basic 6 e roda diretamente no prompt do MS-DOS em ambiente 32bits
Ele se baseia em modulo .bas apenas e pode ser compilado e modificado de acordo com a necessidade
mantendo apenas os direitos autorais dentro dos parâmetros "Project Properties".

Sua estrutura de busca de localização é bem diversificada e totalmente automatizada,
fazendo a leitura do arquivo JSON contendo as informações de previsão do tempo fornecidos pelo wttr.in/
Exemplos de entrada para o prompt:
   -> tempo [cidade] [estado]           --> faz a busca por cidade e estado
   -> tempo [lat],[long]                --> faz a busca por coordenadas (latitude e longitude)
   -> tempo [code]                      --> busca os dados por código de área
   -> tempo                             --> uma vez fazendo uma busca qualquer ele salva a ultima pesquisa
   -> tempo /utc [h]                    --> define o fuso horário para exibir corretamente os horários
   -> tempo /help                       --> cria um guia de ajuda instantâneo na tela
