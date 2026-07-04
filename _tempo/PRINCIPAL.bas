Attribute VB_Name = "PRINCIPAL"
Option Explicit

Private Declare Function AttachConsole Lib "kernel32" (ByVal dwProcessId As Long) As Long
Private Declare Function FreeConsole Lib "kernel32" () As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function Getch Lib "msvcrt" Alias "_getch" () As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const VK_RETURN As Byte = &HD
Private Const KEYEVENTF_KEYUP As Long = &H2
Private Const STD_OUTPUT_HANDLE As Long = -11
Private Const ATTACH_PARENT_PROCESS As Long = -1
Private Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Private Declare Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Private Declare Function AllocConsole Lib "kernel32" () As Long
Private Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Private Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToLong As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Public Const COR_AZUL As Long = &H1
Public Const COR_VERDE As Long = &H2
Public Const COR_VERMELHO As Long = &H4
Public Const COR_AMARELO As Long = &HE
Public Const COR_INTENSIDADE As Long = &H8
Public Const COR_CIANO As Long = COR_VERDE Or COR_AZUL
Public Const COR_ROXO As Long = COR_VERMELHO Or COR_AZUL
Public Const COR_MARROM As Long = COR_VERMELHO Or COR_VERDE
Public Const COR_CINZA_CLARO As Long = COR_VERMELHO Or COR_VERDE Or COR_AZUL
Public Const COR_BRANCO As Long = COR_VERMELHO Or COR_VERDE Or COR_AZUL
Public Const COR_CINZA_ESCURO As Long = COR_INTENSIDADE
Public Const COR_AZUL_VIVO As Long = COR_AZUL Or COR_INTENSIDADE
Public Const COR_VERDE_VIVO As Long = COR_VERDE Or COR_INTENSIDADE
Public Const COR_CIANO_VIVO As Long = COR_VERDE Or COR_AZUL Or COR_INTENSIDADE
Public Const COR_VERMELHO_VIVO As Long = COR_VERMELHO Or COR_INTENSIDADE
Public Const COR_AMARELO_VIVO As Long = &HE
Public Const COR_ROSA As Long = COR_VERMELHO Or COR_AZUL Or COR_INTENSIDADE
Public Const COR_BRANCO_PURO As Long = COR_VERMELHO Or COR_VERDE Or COR_AZUL Or COR_INTENSIDADE
Dim hStdout As Long

Sub Main()
    If AttachConsole(ATTACH_PARENT_PROCESS) = 0 Then
        Exit Sub
    End If
    Dim linhaComando As String
    Dim listaArgs() As String
    Dim primeiroArg As String
    Dim CidadeBusca As String
    
    Dim PegaLocal As String
    Dim PegaUTC As String
    PegaLocal = LerINI("tempo.exe", "local", App.Path & "\config.ini")
    PegaUTC = LerINI("tempo.exe", "utc", App.Path & "\config.ini")
    If PegaUTC = "" Then PegaUTC = "-3"
    
    hStdout = GetStdHandle(STD_OUTPUT_HANDLE)
    ConsoleWrite vbCrLf, COR_BRANCO
    linhaComando = Trim$(Command$)
    Dim argumentos As String
    Dim cidadeFormatada As String
    Dim urlAPI As String
    argumentos = Trim(Command$)
    If Len(linhaComando) > 0 Then
        listaArgs = Split(linhaComando, " ")
        primeiroArg = LCase$(Trim$(listaArgs(0)))
        If primeiroArg = "/help" Or primeiroArg = "-h" Or primeiroArg = "/?" Then
            ExibirAjudaGWeather
            FinalizarConsole
            Exit Sub
        End If
        If primeiroArg = "/utc" Then
            PegaUTC = (Trim$(listaArgs(1))) * 1
            GravarINI "tempo.exe", "utc", PegaUTC, App.Path & "\config.ini"
            ConsoleWrite "-> UTC Atualizado com sucesso: Agora definido para " & PegaUTC & "h" & vbCrLf, COR_VERMELHO_VIVO
            linhaComando = ""
        End If
        GravarINI "tempo.exe", "local", LCase(linhaComando), App.Path & "\config.ini"
        cidadeFormatada = Replace(LCase(linhaComando), " ", "+")
        urlAPI = cidadeFormatada
    Else
      If PegaLocal <> "" Then
        cidadeFormatada = Replace(LCase(PegaLocal), " ", "+")
        urlAPI = cidadeFormatada
      End If
    End If
    
    ConsultarClimaWttr urlAPI, PegaUTC * 1
    FinalizarConsole
End Sub

Public Sub ConsultarClimaWttr(ByVal Localidade As String, ByVal HoraUTC As Integer)
    Dim http As New MSXML2.ServerXMLHTTP60
    Dim sc As New ScriptControl
    Dim url As String
    Dim jsonResponse As String
    Dim urlAPI As String
    Dim RespostaJSON As String
    Dim cidadeFormatada As String
    Dim LonLat As String
    LonLat = "-29.6686,-53.1489"
    If Localidade <> "" Then LonLat = Localidade

    
    ConsoleWrite "-> Baixando dados atualizados via wttr.in... " & vbCrLf, COR_ROXO
    url = "https://wttr.in/" & LonLat & "?format=j1"
    On Error Resume Next
    http.Open "GET", url, False
    http.send
    ConsoleWrite "================================================================================" & vbCrLf, COR_AZUL_VIVO
    ConsoleWrite "           PREVISÃO DO TEMPO MS-DOS - ELETROGB SOFTWARES (2026)" & vbCrLf, COR_VERDE_VIVO
    ConsoleWrite "================================================================================" & vbCrLf, COR_AZUL_VIVO
    If Err.Number <> 0 Then
        ConsoleWrite "-> Erro Crítico: Falha de conexão com o servidor meteorológico!" & vbCrLf, COR_VERMELHO_VIVO
        ConsoleWrite "================================================================================" & vbCrLf, COR_AZUL_VIVO
        Err.Clear: Set http = Nothing: Exit Sub
    End If
    jsonResponse = http.responseText
    sc.Language = "JScript"
    sc.ExecuteStatement "var dados = " & jsonResponse & ";"
    
    Dim WCode As Integer
    Dim Situacao As String
    Dim TempReal As String
    Dim UmidadeReal As String
    Dim CodeClima As String
    Dim SensacaoTermica As String
    Dim VentoVelocidade As String
    Dim PressaoBAR As String
    Dim Precipitacao As String
    Dim UVindice As Integer
    Dim UVindiceTXT As String
    Dim Visibilidade As String
    Dim VentoDirecao As String
    Dim Localizacao As String
    Dim LocEstado As String
    Dim Nuvens As Integer
    Dim LuaIndice As Integer
    Dim LuaFase As String
    Dim LuaNascente As String
    Dim LuaPoente As String
    Dim SolNascente As String
    Dim SolPoente As String
    Dim HoraAtualizada As String
    Dim Latitude As String
    Dim Longitude As String
    
    Localizacao = sc.Eval("dados.nearest_area[0].areaName[0].value")
    LocEstado = sc.Eval("dados.nearest_area[0].region[0].value")
    WCode = sc.Eval("dados.current_condition[0].weatherCode")
    Situacao = sc.Eval("dados.current_condition[0].lang_pt[0].value")
    TempReal = sc.Eval("dados.current_condition[0].temp_C")
    SensacaoTermica = sc.Eval("dados.current_condition[0].FeelsLikeC")
    UmidadeReal = sc.Eval("dados.current_condition[0].humidity")
    PressaoBAR = sc.Eval("dados.current_condition[0].pressure")
    Precipitacao = sc.Eval("dados.current_condition[0].precipMM")
    UVindice = sc.Eval("dados.current_condition[0].uvIndex")
    UVindiceTXT = UVSelect(UVindice)
    VentoVelocidade = sc.Eval("dados.current_condition[0].windspeedKmph")
    VentoDirecao = sc.Eval("dados.current_condition[0].winddir16Point")
    Visibilidade = sc.Eval("dados.current_condition[0].visibility")
    Nuvens = sc.Eval("dados.current_condition[0].cloudcover")
    LuaIndice = sc.Eval("dados.weather[0].astronomy[0].moon_illumination")
    LuaFase = sc.Eval("dados.weather[0].astronomy[0].moon_phase")
    LuaFase = LuaSelect(LuaFase)
    LuaNascente = sc.Eval("dados.weather[0].astronomy[0].moonrise")
    LuaPoente = sc.Eval("dados.weather[0].astronomy[0].moonset")
    SolNascente = sc.Eval("dados.weather[0].astronomy[0].sunrise")
    SolPoente = sc.Eval("dados.weather[0].astronomy[0].sunset")
    HoraAtualizada = sc.Eval("dados.current_condition[0].observation_time")
    Latitude = sc.Eval("dados.nearest_area[0].latitude")
    Longitude = sc.Eval("dados.nearest_area[0].longitude")
    
    ConsoleWrite "   Cidade/localização:                  ", COR_BRANCO
    ConsoleWrite Localizacao & vbCrLf, COR_BRANCO
    ConsoleWrite "   Estado:                              " & LocEstado & vbCrLf, COR_BRANCO
    ConsoleWrite "   Predominância:                       ", COR_BRANCO
    ConsoleWrite GetWeatherDesc(WCode) & vbCrLf, COR_BRANCO
    ConsoleWrite "   Temperatura Agora:                   ", COR_BRANCO
    ConsoleWrite Replace(TempReal, ".", ",") & Chr$(176) & "C" & vbCrLf, COR_VERMELHO_VIVO
    ConsoleWrite "   Sensação Térmica:                    ", COR_BRANCO
    ConsoleWrite Replace(SensacaoTermica, ".", ",") & Chr$(176) & "C" & vbCrLf, COR_VERMELHO_VIVO
    ConsoleWrite "   Umidade do Ar:                       ", COR_BRANCO
    ConsoleWrite UmidadeReal & "%" & vbCrLf, COR_VERDE_VIVO
    ConsoleWrite "   Nebulosidade:                        ", COR_BRANCO
    ConsoleWrite Nuvens & "%" & vbCrLf, COR_BRANCO
    ConsoleWrite "   Velocidade/direção vento:            ", COR_BRANCO
    ConsoleWrite Replace(VentoVelocidade, ".", ",") & " km/h" & " - ", COR_BRANCO
    ConsoleWrite TraduzirDirecaoVento(VentoDirecao) & "" & " [" & VentoDirecao & "]" & vbCrLf, COR_BRANCO
    ConsoleWrite "   Precipitação:                        ", COR_BRANCO
    ConsoleWrite Replace(Precipitacao, ".", ",") & "mm" & vbCrLf, COR_BRANCO
    ConsoleWrite "   Índice Ultravioleta:                 ", COR_BRANCO
    ConsoleWrite UVindice & " [" & UVindiceTXT & "]" & vbCrLf, COR_BRANCO
    ConsoleWrite "   Pressão atmosférica:                 ", COR_BRANCO
    ConsoleWrite Replace(PressaoBAR, ".", ",") & " hPa" & vbCrLf, COR_BRANCO
    ConsoleWrite "   Condição atmosférica:                ", COR_BRANCO
    ConsoleWrite AnalisarPressao(PressaoBAR) & "" & vbCrLf, COR_BRANCO
    ConsoleWrite "================================================================================" & vbCrLf, COR_AZUL_VIVO
    ConsoleWrite "   Horário do nascer/pôr do Sol:        ", COR_BRANCO
    ConsoleWrite Format(CDate(SolNascente), "HH:nn") & " - ", COR_BRANCO
    ConsoleWrite Format(CDate(SolPoente), "HH:nn") & vbCrLf, COR_BRANCO
    ConsoleWrite "   Horário do nascer/pôr da Lua:        ", COR_BRANCO
    ConsoleWrite Format(CDate(LuaNascente), "HH:nn") & " - ", COR_BRANCO
    ConsoleWrite Format(CDate(LuaPoente), "HH:nn") & "" & vbCrLf, COR_BRANCO
    ConsoleWrite "   Fase da lua:                         ", COR_BRANCO
    ConsoleWrite LuaFase & " " & vbCrLf, COR_BRANCO
    ConsoleWrite "   Luminosidade da lua:                 ", COR_BRANCO
    ConsoleWrite LuaIndice & "%" & vbCrLf, COR_BRANCO
    ConsoleWrite "   Última leitura meteorológica:        ", COR_BRANCO
    ConsoleWrite ConverteHoraUTC(HoraAtualizada, HoraUTC) & " [" & HoraUTC & "h UTC]" & vbCrLf, COR_BRANCO
    ConsoleWrite "   Latitude/Longitude:                  ", COR_BRANCO
    ConsoleWrite "" & Latitude & ", " & Longitude & "" & vbCrLf, COR_BRANCO
    ConsoleWrite "   Url/json: ", COR_BRANCO
    ConsoleWrite "" & url & vbCrLf, COR_ROXO
    ConsoleWrite "================================================================================", COR_AZUL_VIVO
End Sub

Function ConverteHoraUTC(HoraUTC As String, UTC As Integer) As String
  Dim dataConvertida As Date
  dataConvertida = CDate(HoraUTC)
  dataConvertida = DateAdd("h", UTC, dataConvertida)
  ConverteHoraUTC = Format(dataConvertida, "HH:nn")
End Function

Public Function TraduzirDirecaoVento(ByVal siglaIngles As String) As String
    Select Case Trim(UCase(siglaIngles))
        Case "N":   TraduzirDirecaoVento = "Norte"
        Case "NNE": TraduzirDirecaoVento = "Nor-nordeste"
        Case "NE":  TraduzirDirecaoVento = "Nordeste"
        Case "ENE": TraduzirDirecaoVento = "Leste-nordeste"
        Case "E":   TraduzirDirecaoVento = "Leste"
        Case "ESE": TraduzirDirecaoVento = "Leste-sudeste"
        Case "SE":  TraduzirDirecaoVento = "Sudeste"
        Case "SSE": TraduzirDirecaoVento = "Sul-sudeste"
        Case "S":   TraduzirDirecaoVento = "Sul"
        Case "SSW": TraduzirDirecaoVento = "Sul-sudoeste"
        Case "SW":  TraduzirDirecaoVento = "Sudoeste"
        Case "WSW": TraduzirDirecaoVento = "Oeste-sudoeste"
        Case "W":   TraduzirDirecaoVento = "Oeste"
        Case "WNW": TraduzirDirecaoVento = "Oeste-noroeste"
        Case "NW":  TraduzirDirecaoVento = "Noroeste"
        Case "NNW": TraduzirDirecaoVento = "Nor-noroeste"
        Case Else:  TraduzirDirecaoVento = "Indisponível"
    End Select
End Function

Public Function AnalisarPressao(ByVal valorPressao As String) As String
    Dim numPressao As Long
    If IsNumeric(valorPressao) Then
        numPressao = CLng(valorPressao)
        Select Case numPressao
            Case Is < 990:      AnalisarPressao = "ALERTA: Tempestade Severa / Ciclone"
            Case 990 To 1000:   AnalisarPressao = "ATENÇÃO: Chuva e Ventos Fortes"
            Case 1001 To 1012:  AnalisarPressao = "Instável: Chance de Chuva"
            Case 1013:          AnalisarPressao = "Normal (Estável)"
            Case 1014 To 1020:  AnalisarPressao = "Bom: Céu Aberto"
            Case 1021 To 1030:  AnalisarPressao = "Firme: Tempo Seco"
            Case Is > 1030:     AnalisarPressao = "ATENÇÃO: Bloqueio / Ar Muito Seco"
            Case Else
                AnalisarPressao = "Estável"
        End Select
    Else
        AnalisarPressao = "Dados Inválidos"
    End If
End Function

Public Function LuaSelect(Valor As String) As String
  Select Case Valor
    Case Is = "New Moon":         LuaSelect = "Lua Nova"
    Case Is = "Waxing Crescent":  LuaSelect = "Lua Nova Crescente"
    Case Is = "First Quarter":    LuaSelect = "Quarto Crescente"
    Case Is = "Waxing Gibbous":   LuaSelect = "Lua Gibosa Crescente"
    Case Is = "Full Moon":        LuaSelect = "Lua Cheia"
    Case Is = "Waning Gibbous":   LuaSelect = "Lua Gibosa Minguante"
    Case Is = "Third Quarter":    LuaSelect = "Quarto Minguante"
    Case Is = "Last Quarter":     LuaSelect = "Quarto Minguante"
    Case Is = "Waning Crescent":  LuaSelect = "Lua Minguante"
  End Select
End Function

Public Function UVSelect(Valor As Integer) As String
  If Valor >= 0 And Valor <= 2 Then UVSelect = "Baixo"
  If Valor >= 3 And Valor <= 5 Then UVSelect = "Moderado"
  If Valor >= 6 And Valor <= 7 Then UVSelect = "Alto"
  If Valor >= 8 And Valor <= 10 Then UVSelect = "Muito Alto"
  If Valor >= 11 Then UVSelect = "Extremo"
End Function



Public Function GetWeatherDesc(Valor As Integer) As String
  Select Case Valor
    Case "113": GetWeatherDesc = "Ensolarado / Céu limpo" 'Sunny / Clear
    Case "116": GetWeatherDesc = "Parcialmente nublado" ' Partly Cloudy
    Case "119": GetWeatherDesc = "Nublado" ' Cloudy
    Case "122": GetWeatherDesc = "Nublado" ' Overcast
    Case "143": GetWeatherDesc = "Névoa" ' Mist
    Case "176": GetWeatherDesc = "Chuva isolada" ' Patchy rain nearby
    Case "179": GetWeatherDesc = "Neve irregular" ' Patchy snow nearby
    Case "182": GetWeatherDesc = "Chuva de granizo esporádica" ' Patchy sleet nearby
    Case "185": GetWeatherDesc = "Chuvisco congelante isolados" ' Patchy freezing drizzle nearby
    Case "200": GetWeatherDesc = "Tempestades" ' Thundery outbreaks in nearby
    Case "227": GetWeatherDesc = "Neve com vento" ' Blowing snow
    Case "230": GetWeatherDesc = "Nevasca" ' Blizzard
    Case "248": GetWeatherDesc = "Névoa" ' fog
    Case "260": GetWeatherDesc = "Nevoeiro congelante" ' freezing fog
    Case "263": GetWeatherDesc = "Chuvisco fraco e isolado" ' Patchy light drizzle
    Case "266": GetWeatherDesc = "Chuvisco leve" ' light drizzle
    Case "281": GetWeatherDesc = "Chuvisco congelante" ' freezing drizzle
    Case "284": GetWeatherDesc = "Chuvisco congelante forte" ' Heavy freezing drizzle
    Case "293": GetWeatherDesc = "Chuva fraca e isolada" ' Patchy light rain
    Case "296": GetWeatherDesc = "Chuva fraca" ' light rain
    Case "299": GetWeatherDesc = "Chuva moderada em alguns momentos" ' Moderate rain at times
    Case "302": GetWeatherDesc = "Chuva moderada" ' Moderate rain
    Case "305": GetWeatherDesc = "Chuvas fortes em alguns momentos" ' Heavy rain at times
    Case "308": GetWeatherDesc = "Chuva forte" ' heavy rain
    Case "311": GetWeatherDesc = "Chuva congelante fraca" ' Light freezing rain
    Case "314": GetWeatherDesc = "Chuva congelante moderada ou forte" ' Moderate or heavy freezing rain
    Case "317": GetWeatherDesc = "Chuva congelada fraca" ' light sleet
    Case "320": GetWeatherDesc = "Chuva com gelo moderada ou forte" ' Moderate or heavy sleet
    Case "323": GetWeatherDesc = "Neve fraca e isolada" ' Patchy light snow
    Case "326": GetWeatherDesc = "Neve fraca" ' light snow
    Case "329": GetWeatherDesc = "Neve moderada em áreas isoladas" ' Patchy moderate snow
    Case "332": GetWeatherDesc = "Neve moderada" ' Moderate snow
    Case "335": GetWeatherDesc = "Neve forte e localizada" ' Patchy heavy snow
    Case "338": GetWeatherDesc = "Neve intensa" ' heavy snow
    Case "350": GetWeatherDesc = "Granizo" ' Ice pellets
    Case "353": GetWeatherDesc = "Chuva fraca e passageira" ' Light rain shower
    Case "356": GetWeatherDesc = "Pancada de chuva moderada ou forte" ' Moderate or heavy rain shower
    Case "359": GetWeatherDesc = "Chuva torrencial" ' Torrential rain shower
    Case "362": GetWeatherDesc = "Chuviscos de neve granulada" ' Light sleet showers
    Case "365": GetWeatherDesc = "Pancadas de chuva com granizo, intensidade moderada ou forte" ' Moderate or heavy sleet showers
    Case "368": GetWeatherDesc = "Pancadas com neve fraca" ' Light snow showers
    Case "371": GetWeatherDesc = "Pancadas com neve moderadas ou fortes" ' Moderate or heavy snow showers
    Case "374": GetWeatherDesc = "Chuviscos com granizo" ' Light showers of ice pellets
    Case "377": GetWeatherDesc = "Pancadas moderadas ou fortes com granizo" ' Moderate or heavy showers of ice pellets
    Case "386": GetWeatherDesc = "Chuva fraca e isolada na região, com trovões." ' Patchy light rain in area with thunder
    Case "389": GetWeatherDesc = "Chuva moderada ou forte com trovões" ' Moderate or heavy rain in area with thunder
    Case "392": GetWeatherDesc = "Neve fraca e isolada na região, com trovões." ' Patchy light snow in area with thunder
    Case "395": GetWeatherDesc = "Neve moderada ou forte com trovoadas na região" ' Moderate or heavy snow in area with thunder
  End Select
End Function


' ====================================================================
' AJUDA DO UTILITÁRIO EM 80 COLUNAS
' ====================================================================
Public Sub ExibirAjudaGWeather()
    ConsoleWrite "================================================================================" & vbCrLf, COR_AZUL_VIVO
    ConsoleWrite "                                 tempo.exe - AJUDA                              " & vbCrLf, COR_VERDE_VIVO
    ConsoleWrite "================================================================================" & vbCrLf, COR_AZUL_VIVO
    ConsoleWrite " Uso do comando:" & vbCrLf, COR_ROXO
    ConsoleWrite "   tempo.exe                 -> Previsão do tempo da localidade padrão" & vbCrLf, COR_BRANCO
    ConsoleWrite "   tempo.exe [cidade]        -> Previsão pesquisado por cidade e estado" & vbCrLf, COR_BRANCO
    ConsoleWrite "   tempo.exe [cod]           -> Previsão do tempo por código de área" & vbCrLf, COR_BRANCO
    ConsoleWrite "   tempo.exe [-29.6,-53.1]   -> Previsão do tempo por latitude e longitude" & vbCrLf, COR_BRANCO
    ConsoleWrite "   tempo.exe [muc]           -> Previsão por código do aeroporto [3 letras]" & vbCrLf, COR_BRANCO
    ConsoleWrite "   tempo.exe /help           -> Exibe esta tela de ajuda" & vbCrLf, COR_BRANCO
    ConsoleWrite "   tempo.exe /utc [h]        -> Define o fuso horário. Ex: tempo.exe /utc -3" & vbCrLf, COR_BRANCO
    ConsoleWrite " Notas do software:" & vbCrLf, COR_ROXO
    ConsoleWrite "   -> Quando uma cidade é pesquisada, o software registra e salva," & vbCrLf, COR_BRANCO
    ConsoleWrite "      depois de salvo basta digitar no prompt: tempo.exe [enter]" & vbCrLf, COR_BRANCO
    ConsoleWrite "      que os dados são pesquisados de forma automática." & vbCrLf, COR_BRANCO
    ConsoleWrite " Liçenças e desenvolvimento:" & vbCrLf, COR_ROXO
    ConsoleWrite "   Freware/opensource - www.eletrogb.com.br - Gilberto Gustavo Böck - (c)2026" & vbCrLf, COR_BRANCO
    ConsoleWrite "================================================================================", COR_AZUL_VIVO
End Sub


Public Function ConsoleErrorParameter(Valor As String)
  ConsoleWrite "-> Erro: O parametro " & Valor & " exige um valor em seguida." & vbCrLf, COR_VERMELHO_VIVO
End Function

Public Function ConsoleWrite(ByVal Texto As String, Optional Cor As Long)
  If Cor Then
    MudarCorConsole Cor
  End If
  Dim hStdout As Long
  Dim BytesEscritos As Long
  Dim textoConvertido As String
  textoConvertido = String$(Len(Texto), 0)
  Call CharToOem(Texto, textoConvertido)
  hStdout = GetStdHandle(STD_OUTPUT_HANDLE)
  Call WriteFile(hStdout, textoConvertido, Len(textoConvertido), BytesEscritos, ByVal 0&)
  RestaurarCorPadrao
End Function

Public Function FinalizarConsole()
  ConsoleWrite vbCrLf
  Call FreeConsole
  Call keybd_event(VK_RETURN, 0, 0, 0)
  Call keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)
  RestaurarCorPadrao
  End
End Function

Public Function MudarCorConsole(ByVal CorHex As Long)
  Dim hConsole As Long
  hConsole = GetStdHandle(STD_OUTPUT_HANDLE)
  SetConsoleTextAttribute hConsole, CorHex
End Function

Public Function RestaurarCorPadrao()
  Dim hConsole As Long
  hConsole = GetStdHandle(STD_OUTPUT_HANDLE)
  SetConsoleTextAttribute hConsole, COR_BRANCO
End Function

Public Function EspacarTexto(ByVal Texto As String, ByVal TamanhoFixo As Integer) As String
  If Len(Texto) > TamanhoFixo Then
    EspacarTexto = Left$(Texto, TamanhoFixo - 3) & "..."
  Else
    EspacarTexto = Texto & Space$(TamanhoFixo - Len(Texto))
  End If
End Function

Public Function ObterDiretorioAtual() As String
  Dim ShellObj As Object
  On Error Resume Next
  Set ShellObj = CreateObject("WScript.Shell")
  ObterDiretorioAtual = ShellObj.CurrentDirectory
  Set ShellObj = Nothing
  If Right$(ObterDiretorioAtual, 1) <> "\" Then
    ObterDiretorioAtual = ObterDiretorioAtual & "\"
  End If
End Function


