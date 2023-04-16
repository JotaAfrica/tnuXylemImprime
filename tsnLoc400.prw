#INCLUDE "PROTHEUS.CH"
#INCLUDE "MSOLE.CH"
#INCLUDE "TOTVS.CH"





*-------------------------*
user function tsnLoc400()
*------------------------*
Local aPergs	:= {}
Local cMascara  := AVSX3('FP0_PROJET',6)
Local nTamanho  := AVSX3('FP0_PROJETO',3) + 30 

aAdd(aPergs, {1, "Nro Proposta/Contrato",  Space( AvSx3("FP0_PROJET",3)) , cMascara,"", "FP0", ".T.", nTamanho , .F.})

// 1 - < aParametros > - Vetor com as configurações
// 2 - < cTitle >      - Título da janela
// 3 - < aRet >        - Vetor passador por referencia que contém o retorno dos parâmetros
// 4 - < bOk >         - Code block para validar o botão Ok
// 5 - < aButtons >    - Vetor com mais botões além dos botões de Ok e Cancel
// 6 - < lCentered >   - Centralizar a janela
// 7 - < nPosX >       -  não centralizar janela coordenada X para início
// 8 - < nPosY >       - Se não centralizar janela coordenada Y para início
// 9 - < oDlgWizard >  - Utiliza o objeto da janela ativa
//10 - < cLoad >       - Nome do perfil se caso for carregar
//11 - < lCanSave >    - Salvar os dados informados nos parâmetros por perfil
//12 - < lUserSave >   - Configuração por usuário

bOk:= { || NeValFP0( aRet[1])  }

if paramBox(aPergs, "Informe um intervalo de Pedidos para gerar a planilha", @aRet, /* bOK */  bOK      , /* aButtons */, /* lCentered */,;
                    /* nPosX */, /* nPosY */, /* oDlgWizard */, /* cLoad */, .T. , .T.)
   geraRel( aRet[1])
endif

return


*---------------------------------*
STATIC  Function NeValFP0(cProjet)
*---------------------------------*
local lRet  := .T. 
local cTmp  :=  GetNextAlias()
local cQuery:= " "

FP0->(dbsetorder(1))
if ! FP0->(dbseek(xfilial("FP0")+cProjeto))
   MsgInfo(oemToAnsi("Proposta/Contrato não encontrado."))
   lRet:= .F. 
endif 
return lRet 

*------------------------*
static function geraRel(cPurch1)
*------------------------*
Local cMascara    := "Arquivo dotx|*.dotx"
Local cTitulo     := "Selecione o arquivo de modelo"
Local nMascpadrao := 1  // para quando existirem mais de uma opção de tipos de arquivo(parâmetro cMascara)
Local cDirinicial := ""  
Local lSalvar     := .F. // .T. para gravação, .F. para leitura
Local nOpcoes     := nOR(GETF_LOCALHARD, GETF_NETWORKDRIVE,GETF_RETDIRECTORY)     // Opções de tratamento interno da função  GETF_LOCALHARD (16)	Apresenta a unidade do disco local., GETF_NETWORKDRIVE (32)	Apresenta as unidades da rede (mapeamento), GETF_RETDIRECTORY (128)	Retorna/apresenta um diretório.
Local lArvore     := .T. /*.T. = apresenta o árvore do servidor || .F. = não apresenta*/
Local lKeepCase   := .T. // Guarda o case original

//cGetFile ( [ cMascara], [ cTitulo], [ nMascpadrao], [ cDirinicial], [ lSalvar], [ nOpcoes], [ lArvore], [ lKeepCase] ) 

cDir := cGetFile(cMascara, cTitulo, nMascpadrao,cDirinicial, lSalvar, nOpcoes,lArvore, lKeepCase) 
if(!File(cDir))
   MsgInfo(oemToAnsi("Arquivo modelo não informado!"))
   return 
Endif

cTitulo     := "Selecione a pasta de gravação"
cDirinicial := "c:\temp\"
cDir := cGetFile(cMascara, cTitulo, nMascpadrao,cDirinicial, lSalvar, nOpcoes,lArvore, lKeepCase) 
if(!File(cDir))
   MsgInfo(oemToAnsi("Pasta não informada!"))
   return 
Endif

If MsgYesNo("Confirma a impressão") 
	// ------------------------------------------------------------------------//
	// Area de declaração de variaveis                                         //
	//-------------------------------------------------------------------------//
	Local cNome		:= "João das Coves"
	Local cEmpresa	:= "AC/DC Industrias Ltda"
	Local cDia		:= "10"
	Local cCidade	:= "São Paulo"
	Local cAdm		:= "Pedro AC"
	Local cNumero	:= "120"
	Local cCPF		:= "123.456.789-00"
	Local nValor	:= 2530.25
	Local cExtenso	:= Extenso(nValor)
	Local cContrato	:= "000123456"
	
	// ------------------------------------------------------------------------//
	// Abre as tabelas para consulta                                           //
	//-------------------------------------------------------------------------//

	set softseek off

	// Inicializa o Ole com o MS-Word
	BeginMsOle()
		If (hWord &gt;= "0")
			IncProc("Processando documento...")
			OLE_CloseLink(hWord) //fecha o Link com o Word

			hWord := OLE_CreateLink()

			OLE_NewFile(hWord,cArquivo)
			If nImpress==1
				OLE_SetProperty( hWord, oleWdVisible,   .F. )
				OLE_SetProperty( hWord, oleWdPrintBack, .T. )
			Else
				OLE_SetProperty( hWord, oleWdVisible,   .T. )
				OLE_SetProperty( hWord, oleWdPrintBack, .F. )
			EndIf


			OLE_SaveAsFile(hWord,cPath+"Contrato" + cContrato + ".doc")
			//OLE_SaveFile(hWord)

			OLE_SetDocumentVar(hWord,"cContrato",cContrato)
			OLE_SetDocumentVar(hWord,"cNome"  	,cNome)
			OLE_SetDocumentVar(hWord,"cEmpresa" ,cEmpresa )
			OLE_SetDocumentVar(hWord,"cDia"		,cDia)
			OLE_SetDocumentVar(hWord,"cCidade"	,cCidade)
			OLE_SetDocumentVar(hWord,"cAdm"    	,cAdm)
			OLE_SetDocumentVar(hWord,"cNumero"  ,cNumero)
			OLE_SetDocumentVar(hWord,"cCPF"     ,cCPF)
			OLE_SetDocumentVar(hWord,"cValor"   ,Transform(nValor,"@E 9,999,999.99"))
			OLE_SetDocumentVar(hWord,"cExtenso" ,cExtenso)

			//--Atualiza Variaveis
			OLE_UpDateFields(hWord)
			OLE_SaveFile ( hWord )

			IF nImpress==1
				OLE_SetProperty( hWord, '208', .F. )
				OLE_PrintFile( hWord, "ALL",,, 1 )
				OLE_CloseLink( hWord )//fecha o Link com o Word
			else
				Aviso("Atenção", "Alterne para o programa do Ms-Word para visualizar o contrato contrato" + cContrato + ".doc ou clique no botao para fechar.", {"Fechar"})
				OLE_SaveAsFile(hWord,cPath+"Contrato" + cContrato + ".doc")
			Endif
		Endif
		
	EndMsOle()

	OLE_CloseLink( hWord )//fecha o Link com o Word


Return


return





















*-----------------------*
user function fModWord()
*-----------------------*
	// ------------------------------------------------------------------------//
	// Area de declaração de variaveis                                         //
	//-------------------------------------------------------------------------//

	Local cCadastro	:= OemtoAnsi("Integração com MS-Word")
	Local aMensagem	:={}
	Local aBotoes   :={}
	Local nOpca		:= 0
	Local nPos		:= 0

	Private cPerg   :=Padr("FMODWORD",10)

	//-------------------------------------------------------------------
	// Cria/Verifica as perguntas selecionadas
	//-------------------------------------------------------------------

	Pergunte(cPerg,.F.)
	AjustaSX1()

	AADD(aMensagem,OemToAnsi("Esta rotina irá imprimir um contrato, clique no botão PARAM para informar o PC a") )
	AADD(aMensagem,OemToAnsi("ser impressa e os demais parametros.") )

	AADD(aBotoes, { 5,.T.,{||  Pergunte(cPerg,.T. )}})
	AADD(aBotoes, { 6,.T.,{|o| nOpca := 1,FechaBatch()}})
	AADD(aBotoes, { 2,.T.,{|o| FechaBatch() }} )

	FormBatch( cCadastro, aMensagem, aBotoes )

	/*
	+------------------------------------------------------------------
	| Variaveis utilizadas para parametros
	+------------------------------------------------------------------
	| Variaveis utilizadas para parametros
	| mv_par01		// Arquivo Modelo
	| mv_par02		// Pasta de destino do documento
	| mv_par03		// Método de saida 1 = Impressora 2 = apenas Arquivo
	+-------------------------------------------------------------------
	*/
	If nOpca == 1
		PRIVATE hWord	 	:= OLE_CreateLink()
		PRIVATE cArquivo 	:= Alltrim(mv_par01)
		PRIVATE cPath    	:= AllTrim(mv_par02)
		PRIVATE nImpress	:= mv_par03

		nPos := Rat("\",cPath)
		If nPos &lt;= 0
			cPath := cPath + "\"
		EndIF

		if(!File(cArquivo))
			Alert("Arquivo modelo não existe!")
			return
		Endif

		IF  Upper( Subst( AllTrim( cArquivo ), - 3 ) ) != Upper( AllTrim( "DOT" ) ) .AND.;
			Upper( Subst( AllTrim( cArquivo ), - 4 ) ) != Upper( AllTrim( "DOTM" ) ) .AND.;
			Upper( Subst( AllTrim( cArquivo ), - 4 ) ) != Upper( AllTrim( "DOTX" ) ) 
	    	MsgAlert( "Arquivo Invalido!"+CRLF+"Extensões permitidas: DOT ou DOTM ou DOTX" )
	    	Return
	    EndIf

		If (hWord &lt; "0")
			Alert("MS-WORD nao encontrado nessa maquina!!!")
			Return
		Endif

		Processa({|| Imprimir() },"Aguarde...")
	EndIf


Return

*-------------------------*
Static Function Imprimir()
*-------------------------*
	// ------------------------------------------------------------------------//
	// Area de declaração de variaveis                                         //
	//-------------------------------------------------------------------------//
	Local cNome		:= "João das Coves"
	Local cEmpresa	:= "AC/DC Industrias Ltda"
	Local cDia		:= "10"
	Local cCidade	:= "São Paulo"
	Local cAdm		:= "Pedro AC"
	Local cNumero	:= "120"
	Local cCPF		:= "123.456.789-00"
	Local nValor	:= 2530.25
	Local cExtenso	:= Extenso(nValor)
	Local cContrato	:= "000123456"
	
	// ------------------------------------------------------------------------//
	// Abre as tabelas para consulta                                           //
	//-------------------------------------------------------------------------//

	set softseek off

	// Inicializa o Ole com o MS-Word
	BeginMsOle()
	IncProc("Processando documento...")
	OLE_CloseLink(hWord) //fecha o Link com o Word

	hWord := OLE_CreateLink()

	OLE_NewFile(hWord,cArquivo)
	If nImpress==1
	   OLE_SetProperty( hWord, oleWdVisible,   .F. )
	   OLE_SetProperty( hWord, oleWdPrintBack, .T. )
	Else
	   OLE_SetProperty( hWord, oleWdVisible,   .T. )
	   OLE_SetProperty( hWord, oleWdPrintBack, .F. )
	EndIf
			OLE_SaveAsFile(hWord,cPath+"Contrato" + cContrato + ".doc")
			//OLE_SaveFile(hWord)

			OLE_SetDocumentVar(hWord,"cContrato",cContrato)
			OLE_SetDocumentVar(hWord,"cNome"  	,cNome)
			OLE_SetDocumentVar(hWord,"cEmpresa" ,cEmpresa )
			OLE_SetDocumentVar(hWord,"cDia"		,cDia)
			OLE_SetDocumentVar(hWord,"cCidade"	,cCidade)
			OLE_SetDocumentVar(hWord,"cAdm"    	,cAdm)
			OLE_SetDocumentVar(hWord,"cNumero"  ,cNumero)
			OLE_SetDocumentVar(hWord,"cCPF"     ,cCPF)
			OLE_SetDocumentVar(hWord,"cValor"   ,Transform(nValor,"@E 9,999,999.99"))
			OLE_SetDocumentVar(hWord,"cExtenso" ,cExtenso)

			//--Atualiza Variaveis
			OLE_UpDateFields(hWord)
			OLE_SaveFile ( hWord )

			IF nImpress==1
				OLE_SetProperty( hWord, '208', .F. )
				OLE_PrintFile( hWord, "ALL",,, 1 )
				OLE_CloseLink( hWord )//fecha o Link com o Word
			else
				Aviso("Atenção", "Alterne para o programa do Ms-Word para visualizar o contrato contrato" + cContrato + ".doc ou clique no botao para fechar.", {"Fechar"})
				OLE_SaveAsFile(hWord,cPath+"Contrato" + cContrato + ".doc")
			Endif
		
	EndMsOle()

	OLE_CloseLink( hWord )//fecha o Link com o Word

Return


Static Function AjustaSx1()
	PutSx1(cPerg,"01","Arquivo Modelo ","","","mv_ch1","C",99,0,0,"G","","DIR","","","mv_par01")
	PutSx1(cPerg,"02","Pasta Destino  ","","","mv_ch2","C",99,0,0,"G","","HSSDIR","","","mv_par02")
	PutSX1(cPerg,"03","Saída          ","","","mv_ch3","N",01,0,0,"C","","","","","mv_par03","Impressora", "", "", "","Arquivo")
Return
