#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <Excel.au3>
#include <Date.au3>				; Usar funções de data
#include <AutoItConstants.au3>	; para usar SplashText
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; 2.0.8 20221123 ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
HotKeySet("{PAUSE}", "Pause")
HotKeySet("^{end}", "Terminate")
;;;;;;;;;;;;;;;;;;

Global $Site = "https://pjd.tjgo.jus.br/"	;Tela Inicial
Global $Title = "Processo Judicial - Mozilla Firefox"	; Título da página do Projud
Global $TJGO = "Processo Eletrônico - TJGO - Mozilla Firefox"
Global $Firefox = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
   If Not FileExists("C:\Program Files (x86)\Mozilla Firefox\firefox.exe") Then
	  $Firefox = "C:\Program Files\Mozilla Firefox\firefox.exe"
   EndIf
Global $Consulta = 0 ;1= Normal, 2= Múltiplo, 3= Não encontrado
Global $CtrlF = ""
Global $FF, $Chave, $Passwd, $Teste, $HexColor
Global $hTimer, $g_iSecs, $g_iMins, $g_iHour, $g_sTime

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; EXCEL ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
SplashTextOn("Geppeto", "by L.A. Tsuruda", 150, 50)
Sleep(1000)
SplashOff()
SplashTextOn("Aguarde", "Iniciando...", 150, 50)
Local $oExcel = _Excel_Open()	; Conecta ao Microsoft Excel
   If @error Then Exit MsgBox("", "_Excel_Open", "Erro ao executar." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
Opt('WinTitleMatchMode', 1)	; Procura pelo nome da janela/ título desejado
;~ 																														Local $Xls = WinGetTitle("Projeto")		; Pega o nome da pasta excel com nome Projeto
																														Local $Xls = WinGetTitle("DOC Gep")		; Pega o nome da pasta excel com nome DOC Geppeto
Local $iPosition = StringInStr($Xls, ".xlsm")+4	; Pega a raiz do nome da pasta
Local $sWorkbook = StringLeft($Xls,$iPosition)	; Pega os caracteres à esquerda, de acordo com a contagem acima para verificar o nome desejado
Local $oWorkbook = _Excel_BookAttach($sWorkbook, "filename")
   If @error Then Exit MsgBox($MB_SYSTEMMODAL, "_Excel_BookAttach", "Erro ao vincular '" & $sWorkbook & "'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
Local $Pasta = "Pastas"
Local $Linha = 6
Local $Cell, $Pdf
Global $Mov = _Excel_RangeRead($oWorkbook,"Validação","f1")
Global $Arq = _Excel_RangeRead($oWorkbook,"Validação","f2")

;~ MsgBox("","",$Mov&"."&$Arq, 2)
While $Mov = 0
   Sleep(250)
   $Mov = _Excel_RangeRead($oWorkbook,"Validação","f1")	; Lê o Tipo de Movimentação selecionada
WEnd

While $Arq = 0
Sleep(250)
$Arq = _Excel_RangeRead($oWorkbook,"Validação","f2")	; Lê o Tipo de Arquivo selecionado
WEnd
SplashOff()	; Fecha tela "Aguardando"
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
WinMinimizeAll()	; Minimiza todas as telas

; Confere se Firefox não estiver executando, e chama o programa
If ProcessExists("Firefox.exe") = False Then
   Run($Firefox,"",@SW_SHOWNORMAL)
   ProcessWait("firefox.Exe")
EndIf

Opt("WinTitleMatchMode", 2)

$FF = WinGetHandle("Firefox")
WinSetState($FF, "", @SW_RESTORE)
Send("{ctrldown}0{ctrlup}")	; restaura zoom da página para 100%

   Login()	; Usuário e Senha
;~    _Excel_RangeWrite($oWorkbook, $oWorkbook.activesheet, $Chave, "G1")	; Registra procurador em G1
   Iniciar()

WinWait($TJGO)	; Aguarda após logon
Send("{shiftdown}{Altdown}8{shiftup}{Altup}")	; Colocar fonte tamanho normal
Sleep(300)
$Teste = PixelSearch(59, 100, 264, 129, 0x1D4875)
   IF @error then send ("{esc}")	; retira popup de Salvar senha
	Verifica_Login()	; Verifica se o nome do Usuário pertence ao CPF informado.

while _Excel_RangeRead($oWorkbook,Default,"A"&$Linha) <> ""	; Protocolo
Send("{shiftdown}{Altdown}8{shiftup}{Altup}")	; Colocar fonte tamanho normal
$hTimer = TimerInit() ; Inicia o Timer e salva em uma variável
SplashTextOn("GepPetO", "Buscando...", 150, 50)

; Verifica se o pdf já foi tratado, e procura pela primeira linha até encontrar campo Vazio em Status
	If _Excel_RangeRead($oWorkbook,Default,"G"&$Linha) <> "" then
		While _Excel_RangeRead($oWorkbook,Default,"G"&$Linha) <> ""
		   $Linha=$Linha+1
		WEnd
		$oExcel.Sheets($Pasta).range("G"&$Linha).Select ;Seleciona a linha após contagem
	EndIf

SplashOff()	; Fecha tela "Aguardando"

	  if _Excel_RangeRead($oWorkbook,Default,"A"&$Linha) = "" then	; Se não houver mais Protocolo
		 MsgBox("", "GepPetO","Não há arquivos a serem tratados." & @CRLF & "Programa finalizado." & @CRLF & "Desenvolvido por L. Akemi Tsuruda")
		 ExitLoop
	  EndIf

$Cell = _Excel_RangeRead($oWorkbook,Default,"C"&$Linha)	; Renomeado
;~ 																											   $Pdf = _Excel_RangeRead($oWorkbook,Default,"A"&$Linha)	; Protocolo GEPPETO
																											   $Pdf = _Excel_RangeRead($oWorkbook,Default,"E3")			; Nome DOC
$Teste = 1
	While $Teste <> 0	; confirmar se retornou para Home
		Local $aCoord = PixelSearch(239, 258, 409,297, 0xf5c9c3) ;Pixel vermelho (Pendências de Leitura)
		$Teste = @error
	WEnd
$Teste = 1	; Verifica Cor da barra abaixo do cabeçalho
Sleep(250)
Send("{shiftdown}{Altdown}3{shiftup}{Altup}")	;atalho para box "Pesquisar Processos"
Sleep(250)
Send("{shiftdown}{home}{shiftup}{del}")	; apagar
Sleep(250)
Send($Cell)
Sleep(500)
Send("{tab}{enter}")

Local $y=246
   While $Teste = 1
	  $HexColor = PixelGetColor(1830, $y)
;~ 															MsgBox("","$HexColor 1",hex($HexColor)&","&$y)
		 If hex($HexColor) = "00009ABF" or hex($HexColor) = "00D2D6DE" Then
			; 0x009ABF Pixel azul claro (barra após consulta de processo único)
			; 0xD2D6DE Pixel cinza (barra após consulta de múltiplos processos)
			if hex($HexColor) = "00009ABF" then	; azul
			   $Teste = 0
			ElseIf hex($HexColor) = "00D2D6DE" Then	; cinza
			   Arquivados() ; Verifica se encontra o processo após arrumar os filtros
			   ExitLoop
			EndIf
		 Else
			$y = $y + 1
			if $y > 265 then $y = 246
		 EndIf
   WEnd

   If hex($HexColor) = "00D2D6DE" then	; cinza
	  Plural()	; Múltiplos processos
	  Unique()	; Único processo

   ElseIf hex($HexColor) = "00009ABF" Then	; azul
	  Send("{ctrldown}f{ctrlup}")		; Procurar
	  sleep(300)
	  Send("{Altdown}o{Altup}")
	  Sleep(500)
	  Send("Peticionar{ENTER}")		; Busca por Bloco
	  Sleep(500)
	  Send("{PGUP}")
	  Sleep(1000)
	  ClickVerde() ; Procura por Seleção e clica
	  Unique()	; Único processo

   Elseif hex($HexColor) = "0038D878"  then
	  Send("{shiftdown}{Altdown}h{shiftup}{Altup}")	; Home
	  _Excel_RangeWrite($oWorkbook, $oWorkbook.activesheet, "Erro", "G" & $Linha)
	  $Linha= $Linha + 1
   EndIf
Local $fDiff = TimerDiff($hTimer) ; Diferença de tempo desde o TimerInit
_TicksToTime(Int(TimerDiff($hTimer)), $g_iHour, $g_iMins, $g_iSecs)	; Converte Timer para hora
Local $sTime = $g_sTime ; salva o tempo calculado
$g_sTime = StringFormat("%02i:%02i:%02i", $g_iHour, $g_iMins, $g_iSecs)	; Transforma em 00:00:00
   If $sTime <> $g_sTime Then ControlSetText("Timer", "", "Static1", $sTime)
_Excel_RangeWrite($oWorkbook, $oWorkbook.activesheet, $g_sTime, "E" & $Linha -1)	; Tempo
_Excel_RangeWrite($oWorkbook, $oWorkbook.activesheet, _Now(), "F" & $Linha -1)	; Data Consulta

$Teste = 1
	While $Teste <> 0	; confirmar se retornou para Home
		Local $aCoord = PixelSearch(239, 258, 409,297, 0xf5c9c3) ;Pixel vermelho (Pendências de Leitura)
		$Teste = @error
	WEnd
WEnd

Logout() ; Fecha a página e desloga da sessão
   Sleep(250)
   MsgBox("","GepPetO","GepPetO finalizado." & @CRLF & "Desenvolvido por L. Akemi Tsuruda")
   WinActivate($Xls)
   _Excel_RangeWrite($oWorkbook, $oWorkbook.activesheet, "Pastas", "H5")
   _Excel_Close($oExcel)
   Exit

   ;;;;;;;;;;;;;;;;;;;;;;; FUNÇÕES DE EXECUÇÃO (em ordem alfabética) ;;;;;;;;;;;;;;;;;;;;;;;

Func Acessar()	; Login e Senha
;~ WinSetState($FF, "", @SW_SHOWMAXIMIZED)
;~ WinActivate($FF)
Sleep(300)
Local $Erro = 1
   While $Erro <> 0
	  Local $aCoord = PixelSearch(1087, 497, 1099, 510, 0x333333) ;Bonequinho preto Usuário
	  $Erro = @error
;~ 	  MsgBox("","$Erro",$Erro)
		 if $Erro = 0 Then
			MouseClick("left", $aCoord[0], $aCoord[1], 2,1)
			ExitLoop
		 EndIf
   WEnd
   sleep(500)
   Send("{del}")
   send($Chave)	;Usuário
   Send("{tab}")
   Sleep(200)
   Send("{del}")
   Send($Passwd) ;Senha
   Sleep(200)
   Send("{ENTER}")
   Sleep(1000)

; Validação da senha (se inválida ou dados incorretos)
      Local $aCoord = PixelSearch (1087, 497, 1099, 510, 0xDA3C25)	; Usuário vazio, vermelho
	  if Not @error Then
		 Send("{F5}")	; Atualizar página
		 Sleep(300)
		 MsgBox ("","Atenção!", "Usuário ou Senha digitados incorretamente." & @CRLF & "Re-insira os dados.")
		 Login()	; Usuário e Senha
		 Acessar()
	  EndIf

      Local $aCoord = PixelSearch (1087, 567, 1097, 579, 0xDA3C25)	; Senha vazia, vermelho
	  if Not @error Then
		 Send("{F5}")	; Atualizar página
		 Sleep(300)
		 MsgBox ("","Atenção!", "Usuário ou Senha digitados incorretamente." & @CRLF & "Re-insira os dados.")
		 Login()	; Usuário e Senha
		 Acessar()
	  EndIf

   Local $aCoord = PixelSearch (826, 743, 846, 755, 0xA94442)	;Incorreto, vermelho
	  if Not @error Then
	  Send("{Altdown}d{Altup}")	;Seleciona barra de endereço
	  sleep(300)
	  Send($Site)
	  sleep(250)
	  SEND("{ENTER}")
		 Sleep(300)
		 MsgBox ("","Atenção!", "Usuário ou Senha digitados incorretamente." & @CRLF & "Re-insira os dados.")
		 Login()	; Usuário e Senha
		 Acessar()
	  EndIf

WinWait($TJGO)
$Cn="Sem conexão com o servidor PJD."
sleep(500)	; Teste de conexão
Local $SemCn = PixelGetColor(923, 307)	; botão Fechar
   if hex($SemCn) = "00004A80"  then
	  MsgBox("", "Atenção", $Cn & @CRLF & "Tente novamente mais tarde.")
	  Exit
   EndIf


EndFunc

Func Arquivados() ; Verifica se encontra o processo após arrumar os filtros
Sleep(1000)
MouseClick("left", 629, 340, 1) ; Seleciona opção "Todos"
Sleep(500)
Send("{tab 6}{del}")		; Apaga filtro "Ativo"
Sleep(250)
Send("{tab}")		; Apaga filtro "Ativo"
;~ Send("{tab 4}")		; Apaga filtro "Ativo"
;~ Sleep(500)
;~ Send("{ENTER}")
$Consultar = 1
   While $Consultar <> 0
	  Local $aCoord = PixelSearch(1772, 725, 1862, 747, 0x004a80) ; botão Consultar
	  $Consultar = @error
;~ 	  Mousemove($aCoord[0], $aCoord[1],1)
		 If not @error Then
			MouseClick("left", $aCoord[0], $aCoord[1], 1, 1)
		 EndIf
   WEnd

Local $y=246
   While $Teste = 1
	  $HexColor = PixelGetColor(1830, $y)
;~ 															MsgBox("","$HexColor 1",hex($HexColor)&","&$y)
		 If hex($HexColor) = "00009ABF" or hex($HexColor) = "00D2D6DE" Then
			; 0x009ABF Pixel azul claro (barra após consulta de processo único)
			; 0xD2D6DE Pixel cinza (barra após consulta de múltiplos processos)
			if hex($HexColor) = "00009ABF" then	; azul
			   $Teste = 0
			ElseIf hex($HexColor) = "00D2D6DE" Then	; cinza
			   NRE() ; Verifica se protocolo foi encontrado
			   ExitLoop
			EndIf
		 Else
			$y = $y + 1
			if $y > 265 then $y = 246
		 EndIf
   WEnd
EndFunc

Func Iniciar()
   If WinExists($TJGO)= True then	;Confere se já existe janela PJD aberta
	  WinSetState($FF, "", @SW_SHOWMAXIMIZED)	;Aguarda focar na tela
	  sleep(250)
	  Send("{ctrldown}{Altdown}r{ctrlup}{Altup}")	;atalho para logout

   ElseIf WinExists($Title)= True then
	  WinSetState($FF, "", @SW_SHOWMAXIMIZED)
	  WinActivate($FF)
	  Send("{F5}")	; Atualizar página
	  sleep(250)
	  Send("{ctrldown}0{ctrlup}")	; restaura zoom da página para 100%

	Else	;Se não existir, vai acessar
	  ;Aguarda focar na tela
	  WinSetState($FF, "", @SW_SHOWMAXIMIZED)
	  WinActivate($FF)
	  sleep(500)
	  Send("{ctrldown}t{ctrlup}")	;Nova aba
	  sleep(1000)
	  Send("{Altdown}d{Altup}")	;Seleciona barra de endereço
	  sleep(300)
	  Send($Site)
	  sleep(250)
	  SEND("{ENTER}")
	  Sleep(500)
	  WinActivate($Title)	; Aguarda a página de Logon
	EndIf
Call("Acessar") ; Função para preencher Login e Senha
EndFunc

Func Login()	; Usuário e Senha
   $Chave = _Excel_RangeRead($oWorkbook,Default,"G1")
;~    $Chave=InputBox("Login", "Digite seu Usuário:", "", "", '', '', Default, Default, 0, WinGetHandle(AutoItWinGetTitle()) * WinSetOnTop(AutoItWinGetTitle(), '', 1))
   $Passwd = InputBox("Login", "Digite sua senha:", "", "*", '', '', Default, Default, 0, WinGetHandle(AutoItWinGetTitle()) * WinSetOnTop(AutoItWinGetTitle(), '', 1))
;~    $Chave="36982741168"
;~    $Passwd="Faculdade65@"
    WinActivate($FF)
 EndFunc

 Func Logout() ; Fecha a página e desloga da sessão
   if WinExists($TJGO)= True then ; Se estiver logado
	  WinSetState($FF, "", @SW_SHOWMAXIMIZED)    	  ;Aguarda focar na tela
	  WinActivate($Title)
	  Sleep(1000)
	  Send("{ctrldown}{Altdown}r{ctrlup}{Altup}")	;atalho para logout
	  Sleep(1000)
   EndIf

   WinWaitActive($Title)	; Aguarda a página de Logon
   Send("{ctrldown}w{ctrlup}")	; fecha aba
   WinWaitClose($Title)
EndFunc

func NRE() ; Verifica se protocolo foi encontrado
$Teste = 1
Local $NREx = 721
   While $Teste = 1
   Local $NRE = PixelGetColor(531, $NREx)	; Box "Opções" (seta)
;~ 															   MsgBox("","$NRE",hex($NRE))
	  if hex($NRE) <> "00858585"  then
		 Send("{ctrldown}f{ctrlup}")		; Procurar
		 Sleep(250)
		 Send("Nenhum registro encontrado")		;  Procura expressão
		 Sleep(500)
		 Send("{ENTER}")
		 Sleep(500)
		 $NRE = PixelGetColor(1129, 837)
			if hex($NRE) = "0038D878"  then	; verde
			   Send("{ctrldown}f{ctrlup}{esc}")
			   $HexColor = $NRE
			   ExitLoop
			Else
			   $NREx = $NREx + 1
			   if $NREx > 740 then $NREx = 730
			EndIf
	  ElseIf hex($NRE) = "00858585"  then
		 $Teste = 0
	  EndIf
   WEnd
EndFunc

Func Plural()	; Múltiplos processos
$CtrlF = _Excel_RangeRead($oWorkbook,Default,"B"&$Linha)
   Lupa() ; Procura pelo protocolo na página
   Sleep(250)
   ClickVerde() ; Procura por Seleção Verde e clica
   Sleep(250)
   Send("{tab}{enter}")
endFunc

FUNC Unique()	; Único processo
Opt("SendKeyDelay", 50)
ClipPut("")
;~ 																					   ClipPut(@ScriptDir&"\PDFs\" & $Pdf)	; Caminho do arquivo na Pasta dos pdfs
																					   ClipPut(@ScriptDir&"\Doc\" & $Pdf)	; Doc Geppeto
Sleep(2000)
$Teste = 1	; box Tipo Movim
   While $Teste <> 0
   Local $aCoord = PixelSearch(433, 360, 576, 530, 0xA999A9) ;Pixel púrpura "Informe tipo Movim."
;~    Local $aCoord2 = PixelSearch(433, 360, 576, 530, 0xA899A8) ;Pixel púrpura "Informe tipo Movim." pc Fábia
   $Teste = @error
	  if Not $Teste = 1 Then
		 MouseClick("left", $aCoord[0], $aCoord[1], 2)
		 ExitLoop
	  EndIf
;~    Sleep(350)
WEnd
Sleep(3000)
;~ Send(_Excel_RangeRead($oWorkbook,Default,"A2")) ; Célula com Tipo de Mov.
   If $Mov = 1 Then
	  Send("juntada -> petição -> Contest")
   ElseIf $Mov = 3 Then
	  Send("juntada -> petição -> recurso inte")
   ElseIf $Mov = 2 Then
	  Send("jun")
	  Sleep(2000)
	  Send("{enter}")
   EndIf
   if $Mov <> 2 then
	  Local $Select = 1	; Seleção Tipo
	  Local $z = 545	; início da coordenada
		 While $Select <> 9
	  ;~ 	  Sleep(250)
			Local $HexColor = PixelGetColor(539, $z)
			   if hex($HexColor) = "005897fb"  then
				  $Select = 9
			   Else
				  $z = $z + 1
				  if $z > 605 then $z = 545
			   EndIf
			WEnd
   EndIf
Send("{down}{enter}")
Sleep(250)
$Teste = 1	; box Tipo Documento
   While $Teste <> 0
   Local $aCoord = PixelSearch(552, 707, 570, 756, 0xA999A9) ;Pixel púrpura "Informe tipo do Docum"
;~    Local $aCoord = PixelSearch(552, 707, 570, 756, 0xA899A8) ;Pixel púrpura "Informe tipo do Docum" pc Fábia
   $Teste = @error
;~ 															MsgBox("","$Teste tipo doc",$Teste)
	  if Not @error Then
		 MouseClick("left", $aCoord[0], $aCoord[1], 1)
		 ExitLoop
	  EndIf
;~    Sleep(350)
WEnd
Sleep(2000)
;~ Send(_Excel_RangeRead($oWorkbook,Default,"B2")) ; Célula com Tipo de Documento
   If $Arq = 1 Then
	  Send("Outros")
   ElseIf $Arq = 2 Then
	  Send("Petição")
   EndIf
Sleep(1000)
$z = 773	; início da coordenada
Local $Select = 1	; Seleção Tipo
   While $Select <> 9
	  Sleep(10)
	  Local $HexColor = PixelGetColor(551, $z)
		 if hex($HexColor) = "005897fb"  then
			$Select = 9
		 Else
			$z = $z + 1
			if $z > 832 then $z = 773
		 EndIf
   WEnd
Sleep(1000)
Send("{enter}")
Sleep(250)
Send("{tab}")
Sleep(250)
Send("{shiftdown}{Altdown}o{shiftup}{Altup}")	; Alcançar botão "Selecionar documento"
Send("{enter}")
sleep(500)
;~ 													  MouseMove(1438, 935, 5) 												; Ativar na v. Edilson
WinWait("[CLASS:#32770]")
WinWaitActive("[CLASS:#32770]")
sleep(350)
;~ 													  ControlClick("Enviar arquivo","","[CLASS:Edit; INSTANCE:1]")			; Ativar v. Rodrigo
Send("{ctrldown}v{ctrlup}")																								; Desativar Rodrigo
;~ 													  Send(clipget())														; Substituir ctrl v por este. Versão Rodrigo
sleep(750)																													; Aumentar para 1500. Versão Rodrigo
;~ Send("{Altdown}a{Altup}")		; cola
WinActivate("[CLASS:#32770]", "")
ControlClick("Enviar arquivo","&Abrir","[CLASS:Button; INSTANCE:1]")
Sleep(1000)
$Teste = 1
   While $Teste <> 0
	  Local $aCoord = PixelSearch(557, 804, 569, 897, 0x444444) ;Seta botão Excluir
	  $Teste = @error
   WEnd
;~ 																  MsgBox("","$Teste",@error)
Sleep(1000)
$Teste = 1
   While $Teste = 1
	  Local $aCoord = PixelSearch(1773, 889, 1787, 931, 0x004A80) ; botão INCLUIR
	  $Teste = @error
	  MouseWheel("down", 5)
;~ 																  MsgBox("","INCLUIR",@error)
	  If $Teste = 0 Then
;~ 																  MsgBox("","If not @error",@error)
			MsgBox("","INCLUIR",$aCoord[0]&"-"& $aCoord[1], 1)
			Sleep(500)
			MouseClick("left", $aCoord[0]+5, $aCoord[1]+5, 1, 1)
			$HexColor = 0
			   While hex($HexColor) <> "00004A80"	; Pop up de confirmação
				  Local $HexColor = PixelGetColor(1007, 510)
;~ 																  MsgBox("","while 2",hex($HexColor))
				  if Not hex($HexColor) = "00004A80"  then
					 Sleep(500)
					 Send("{shiftdown}{Altdown}i{shiftup}{Altup}")
					 ExitLoop
				  EndIf
			   WEnd
;~ 			MsgBox("","while 1",hex($HexColor))
			ExitLoop
		 EndIf
	  WEnd
Sleep(500)
;~															MsgBox("","Botão Sim",hex($HexColor))
$Teste = 1
   While $Teste <> 0
	  Local $aCoord = PixelSearch(968, 509, 1012, 528, 0x004a80) ; botão Sim
	  $Teste = @error
	  Mousemove($aCoord[0], $aCoord[1],1)
		 If not @error Then
			MouseClick("left", $aCoord[0], $aCoord[1], 1, 1)
			Sleep(250)
			Send("{shiftdown}{Altdown}s{shiftup}{Altup}")
			ExitLoop
		 EndIf
   WEnd
WinWait($TJGO)
MouseWheel("up", 20)
$Verde = 1	; Peticionado com sucesso
	While $Verde <> 0
		Local $HexColor = PixelGetColor(1371, 278)
			if hex($HexColor) = "0000733E"  then
			   Sleep(400)
			   Send("{shiftdown}{Altdown}h{shiftup}{Altup}")	; Home
			   _Excel_RangeWrite($oWorkbook, $oWorkbook.activesheet, "Tratado", "G" & $Linha)
			   $Linha= $Linha + 1
			  ExitLoop
			Else
			   MouseWheel("up", 20)
			EndIf
	WEnd
EndFunc

func Verifica_Login()
$Teste = 1
Global $Nome = _Excel_RangeRead($oWorkbook,"Pastas","I1")	; Nome cadastrado para o CPF
Local $NREx = 190
;~     															   MsgBox("","$Nome",$Nome, 1)
   Send("{ctrldown}f{ctrlup}")		; Procurar
   Sleep(1000)
   Send($Nome)		;  Procura expressão
   Sleep(500)
;~    Send("{ENTER}")
;~    Sleep(500)
WinWait($TJGO)
CorCaixaDeEntrada()	;Verifica se a tela inicial terminou de carregar (confere a cor da letra C)

   While $Teste = 1
   Local $NRE = PixelGetColor(1553, $NREx)	; Box "Opções" (seta)
;~ 															   MsgBox("","$NRE",hex($NRE))
	  if hex($NRE) <> "00858585"  then
		 $NRE = PixelGetColor(1553, $NREx)
			if hex($NRE) = "0038D878"  then	; verde
			   Send("{ctrldown}f{ctrlup}{esc}")
			   $HexColor = $NRE
			   ExitLoop
			Else
			   $NREx = $NREx + 1
			   if $NREx > 246 then
				  $NREx = 190
				  MsgBox("","Falha na autenticação do Geppeto", "Este CPF não consta em nossa base")
				  Terminate()
			   EndIf
			EndIf
	  ElseIf hex($NRE) = "00858585"  then
		 $Teste = 0
	  EndIf
   WEnd

EndFunc

#cs
Func LongWay()	; Código usando menu Processos
   Sleep(250)
   Send("{shiftdown}{Altdown}r{shiftup}{Altup}")	;atalho para Menu Processos
   Sleep(250)
   Local $Teste = 1
	  While $Teste <> 0
		 Local $aCoord = PixelSearch(1845, 939, 1828, 882, 0x8D89AC) ;Pixel roxo
		 $Teste = @error
	  WEnd
   Sleep(250)
	;~    MsgBox("","$Teste",$Teste)
   Send("{tab 6}{enter}")	; anda até "consultar" e seleciona
   $Teste = 1
	  While $Teste <> 0
		 Local $aCoord = PixelSearch(1847, 750, 1840, 693, 0x624a80) ;Pixel violeta no botão Consultar
		 $Teste = @error
	   WEnd
	Sleep(250)
	Send("{tab}{right}")
	Sleep(250)
	Send("{tab 2}" & $Cell)	; Campo para informar número do processo
	Sleep(300)
	Send("{tab 4}")
	Sleep(250)
	Send("{del}{esc}")	; zera campo Situação
	Sleep(250)
	;~ Send("{ctrldown}{Altdown}c{ctrlup}{Altup}")	; Consultar
	Send("{tab 3}{enter}")	; botão "consultar"
EndFunc
#ce

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; FUNÇÕES DE TESTE ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Func ClickVerde() ; Procura por Seleção Verde e clica
$A=1
While $A=1
   Local $aCoord = PixelSearch (1895, 240, 37, 983, 0x38d878)	; Verde (destaque)
		 $A=@error
		 if $A = 0 Then
			Send("{ctrldown}f{ctrlup}{esc}")		; Fechar Procurar
			Sleep(250)
;~ 			MsgBox("","",$aCoord[0]&"-"&$aCoord[1])
;~ 			Mousemove($aCoord[0], $aCoord[1],1)
			MouseClick("left", $aCoord[0], $aCoord[1], 1, 1)
		 EndIf
WEnd
EndFunc

Func CorCaixaDeEntrada()	;Verifica se a tela inicial terminou de carregar (confere a cor da letra C)
   $C=1
While $C=1
   Local $aCoord = PixelSearch (7, 190, 31, 225, 0x333333)	; Preto
		 $C=@error
		 if $C = 0 Then
			Sleep(250)
		 EndIf
	  WEnd
EndFunc

Func Lupa() ; Procura pelo protocolo na página
   ClipPut($CtrlF)
   sleep(1000)
   Send("{ctrldown}f{ctrlup}")		; Atalho Procurar
   Sleep(1000)
   Send("^v{ENTER}")		;  Procura expressão
   ClipPut("")
   $CtrlF=""
EndFunc

Func Vermelho()	; Verifica se operação é NRE
Sleep(250)
   Local $aCoord = PixelSearch (80, 1023, 168, 1026, 0xff6666)	; Vermelho (campo "procurar")
   if Not @error Then
	  Send("{ctrldown}f{ctrlup}{esc}")		; Fechar Procurar
	  $HexColor = "0x38D878"
   EndIf
EndFunc
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; OUTROS ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Func Terminate()

Logout() ; Fecha a página e desloga da sessão
MsgBox("","GepPetO - Desenvolvido por L. Akemi Tsuruda","Geppeto finalizado.")
_Excel_RangeWrite($oWorkbook, $oWorkbook.activesheet, "Finalizado", "H5")
_Excel_Close($oExcel)
    Exit
 EndFunc

 Func Pause()
   MsgBox($mb_systemmodal,"GepPetO - Desenvolvido por L. Akemi Tsuruda","Em pausa." & @CRLF & "Se desejar finalizar, digite [Ctrl + END]")
EndFunc

; Versionamento a partir da versão 2.0:
;~ 2.0.8 20221107 - Aumento do tempo para selecionar o tipo de Movimentação (l. 446 e 450). Aumentada área de busca (l. 480. Aumento do tempo de espera
;~ 					 (l.485) e alterado código de rolagem para baixo para o loop que aguarda o botão Incluir (l. 490).
;~ 2.0.7 20220704 - Incluído "Page Up" após a busca da palavra "Peticionar" devido a erro na tela do PJD que rola para o fim da página quando realiza esta ação. Linhas 145 e 146.
;~ 2.0.6 20220413 - Incluído código de verificação do carregamento da Caixa de Entrada (Func CorCaixaDeEntrada) e colocado na linha 544 na etapa de verificação do usuário.
;~ 2.0.6 DOC 20220303 - Acrescentada função que lê o nome (agora variável, preenchido no excel) do documento a ser inserido pelo DOC, linhas 99 e 382.
;~ 2.0.5 Adaptação do Código nas linhas 395 a 416: para conseguir selecionar o Tipo "Juntada -> Petição", é necessário usar a tecla {Enter} no lugar do posicionamento e acrescentada opção de Recurso Interposto.
;~ 2.0.4 20220209 - Alteração da Juntada Petição Outros para Juntada Petição
;~ 2.0.3 DOC 20220203 - Incluídas linhas para aproveitamento do código para uso do Geppeto para upload de Documento único no PJD (linhas com nome do Exel e endereço do documento a ser carregado).
;~ 2.0.3 20220126 - Alteração no PJD de "juntada d petição" para "juntada -> petição -> ou", linha 365. Alteração de "contestação apresentada" para "juntada -> petição -> contest", linha 363.
;~ 2.0.2 20210910 - Aumentado range de busca dos campos de Tipo de Mov. e Arquivo, pois o campo Classe apresenta texto muito longo que altera layout.
;~ 2.0.1 20210721 - Acrescentado tempo para sumir o box de nome (l.529) e acrescentado sleep na l. 530.
;~ 2.0   20210623 - Nova função Verifica_Login, para identificar se cpf tem seu nome registrado no servidor. Alterado forma de fazer login
;~			 linha 317 substituiu linha 318, pois, para usar a nova função, ela precisa puxar diretamente do Excel o CPF do Usuário.



;~ Versionamento da versão 1.xx:
; Versão 1 Beta 1305 - //revogado pela versão 1.3)// Incluído log de erro de conexão ao PJD (Sem conexão com servidor PJD)
; Versão 1.1 Beta 1306 - //revogado pela versão 1.3)//	Alterado posição do pixel de 257 para 258, na linha 104 (versão ff 45)
; Versão 1.2 Beta 1406 - Coordenadas alteradas para encontrar cor da seta 'excluir', na linha 367
; Versão 1.3 Beta 1906 - Melhorado teste para encontrar a cor da barra (cinza ou azul) abaixo do cabeçaho (linhas 101 a 118)
;~ 					   - Corrigido código de erro de conexão(Sem conexão com servidor PJD, linhas 227 a 234)
; Versão 1.3.1 Beta 2406 - Retornado código de verificação da Home (linhas 147 a 151)
; Versão 1.4 Beta 2606 - Alterados códigos de verificação da Home (verifica cabeçalho 'Pendências de Leitura' em vez da barra vermelha)
;~ 						e alterado posicionamento de teste no retorno do loop de busca (era verificado depois da seleção da caixa de
;~ 						busca, agora está antes da seleção da caixa, para tentar identificar lentidão do carregamento da página.
;~ Versão 1.5 Beta 2009 - Melhorado código de busca dos boxes de pesquisa de Tipo Mov. e Documento (318, 330, 337, 350).
;~ 						Incluído sleep dentro do código de verificação do box de Tipo Mov. (325 para 324, e 352)
;~ Versão 1.6 Beta 2309 - Incluído códigos de loop para procurar seleção dos Tipos (linhas 327 a 337, 353 a 363) e diminuir incidência do 0 (zero).
;~ Versão 1.6.1 Beta 0110 - Encontrada possível causa do erro do 0 (AutoIt não vinculava valor às variáveis de Tipo de Mov. e Doc. Alteradas linhas (35, 36, 326, 352)
;~ Versão 1.7 Beta 0710 - Alterada forma de captação dos Tipos de Mov. e Doc. (24 a 48; 334 a 342; 370 a 374); Incluído Delay para digitação (323)
;~ Versão 1.8 Beta 0512 - Atualizada linha 404, pois houve alteração no layout (mudança na posição da "seta Excluir").
;~ Versão 1.9 Beta 0912 - Atualizada linha 292, pois houve alteração no layout (mudança na posição da "seta Opções").
;~ Versão 1.9.1 Beta 1312 - Atualizada Func NRE (291, 292, 307 a 310), pois foi identificado diferença de pixels no de algumas máquinas layout (mudança na posição da "seta Opções").
;~ Versão 1.9.2 Beta 2901 - Atualizada Func NRE (302), pois foi identificado diferença de pixels no layout.
;~ Versão 1.9.3 Beta 3101 - Inserido código de tratamento para operações em segredo de justiça (linha 496).
;~ Versão 1.9.4 Beta 2503 - Alterado código para clicar no botão "Abrir", para selecionar documento (linhas 404 e 405)
;~ Versão 1.9.5 Beta 2505 - Várias alterações no código para corrigir problema de filtro interno do PJD (somente procurava por protocolos ativos e próprios).
;~ 						    Incluída função "Arquivados" (249 a 287) que retira os filtros e refaz a busca. Alterada a linha 121 de "NRE()" para "Arquivados()".
;~ Versão 1.9.6 1408 - Nova função Logout para garantir que foi deslogado (323 a 335). Corrigido "enter" para busca de função (incluído Tab), linha 109. Modificado código de finalização, 164 a 169.
;~ Versão 1.9.7 0903 - Alterados os pixels para busca dos campos de "Tipo de Arquivo", "Movimento", "Seta Excluir" e botão "Incluir".
;~ Versão 1.9.8 3105 - Mudança no Layout do PJD: alteração nas coordenadas da linha 411 (Tipo de Documento) e para encontrar cor azul (para confirmar o tipo), linhas 427 a 436.
;~ 	  				 - L.464: Adicionado MouseWheel para baixo. L. 468 atualizadas coordenadas p/ btn Incluir, l.475 adicionado "+5" para melhorar clique. L.505: Adicionada MouseWhheel para cima.
;~ 	      1.9.83 1006	- Alteração da coordenada para processos múltiplos poderem encontrar o campo de Tipo de Movimento.
