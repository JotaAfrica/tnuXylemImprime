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

// 1 - < aParametros > - Vetor com as configura��es
// 2 - < cTitle >      - T�tulo da janela
// 3 - < aRet >        - Vetor passador por referencia que cont�m o retorno dos par�metros
// 4 - < bOk >         - Code block para validar o bot�o Ok
// 5 - < aButtons >    - Vetor com mais bot�es al�m dos bot�es de Ok e Cancel
// 6 - < lCentered >   - Centralizar a janela
// 7 - < nPosX >       -  n�o centralizar janela coordenada X para in�cio
// 8 - < nPosY >       - Se n�o centralizar janela coordenada Y para in�cio
// 9 - < oDlgWizard >  - Utiliza o objeto da janela ativa
//10 - < cLoad >       - Nome do perfil se caso for carregar
//11 - < lCanSave >    - Salvar os dados informados nos par�metros por perfil
//12 - < lUserSave >   - Configura��o por usu�rio

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
   MsgInfo(oemToAnsi("Proposta/Contrato n�o encontrado."))
   lRet:= .F. 
endif 
return lRet 

*------------------------*
static function geraRel(cPurch1)
*------------------------*
Local cMascara    := "Arquivo dotx|*.dotx"
Local cTitulo     := "Selecione o arquivo de modelo"
Local nMascpadrao := 1  // para quando existirem mais de uma op��o de tipos de arquivo(par�metro cMascara)
Local cDirinicial := ""  
Local lSalvar     := .F. // .T. para grava��o, .F. para leitura
Local nOpcoes     := nOR(GETF_LOCALHARD, GETF_NETWORKDRIVE,GETF_RETDIRECTORY)     // Op��es de tratamento interno da fun��o  GETF_LOCALHARD (16)	Apresenta a unidade do disco local., GETF_NETWORKDRIVE (32)	Apresenta as unidades da rede (mapeamento), GETF_RETDIRECTORY (128)	Retorna/apresenta um diret�rio.
Local lArvore     := .T. /*.T. = apresenta o �rvore do servidor || .F. = n�o apresenta*/
Local lKeepCase   := .T. // Guarda o case original

//cGetFile ( [ cMascara], [ cTitulo], [ nMascpadrao], [ cDirinicial], [ lSalvar], [ nOpcoes], [ lArvore], [ lKeepCase] ) 

cDir := cGetFile(cMascara, cTitulo, nMascpadrao,cDirinicial, lSalvar, nOpcoes,lArvore, lKeepCase) 
if(!File(cDir))
   MsgInfo(oemToAnsi("Arquivo modelo n�o informado!"))
   return 
Endif

cTitulo     := "Selecione a pasta de grava��o"
cDirinicial := "c:\temp\"
cDir := cGetFile(cMascara, cTitulo, nMascpadrao,cDirinicial, lSalvar, nOpcoes,lArvore, lKeepCase) 
if(!File(cDir))
   MsgInfo(oemToAnsi("Pasta n�o informada!"))
   return 
Endif

If MsgYesNo("Confirma a impress�o") 
	// ------------------------------------------------------------------------//
	// Area de declara��o de variaveis                                         //
	//-------------------------------------------------------------------------//
	Local cNome		:= "Jo�o das Coves"
	Local cEmpresa	:= "AC/DC Industrias Ltda"
	Local cDia		:= "10"
	Local cCidade	:= "S�o Paulo"
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
				Aviso("Aten��o", "Alterne para o programa do Ms-Word para visualizar o contrato contrato" + cContrato + ".doc ou clique no botao para fechar.", {"Fechar"})
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
	// Area de declara��o de variaveis                                         //
	//-------------------------------------------------------------------------//

	Local cCadastro	:= OemtoAnsi("Integra��o com MS-Word")
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

	AADD(aMensagem,OemToAnsi("Esta rotina ir� imprimir um contrato, clique no bot�o PARAM para informar o PC a") )
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
	| mv_par03		// M�todo de saida 1 = Impressora 2 = apenas Arquivo
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
			Alert("Arquivo modelo n�o existe!")
			return
		Endif

		IF  Upper( Subst( AllTrim( cArquivo ), - 3 ) ) != Upper( AllTrim( "DOT" ) ) .AND.;
			Upper( Subst( AllTrim( cArquivo ), - 4 ) ) != Upper( AllTrim( "DOTM" ) ) .AND.;
			Upper( Subst( AllTrim( cArquivo ), - 4 ) ) != Upper( AllTrim( "DOTX" ) ) 
	    	MsgAlert( "Arquivo Invalido!"+CRLF+"Extens�es permitidas: DOT ou DOTM ou DOTX" )
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
	// Area de declara��o de variaveis                                         //
	//-------------------------------------------------------------------------//
	Local cNome		:= "Jo�o das Coves"
	Local cEmpresa	:= "AC/DC Industrias Ltda"
	Local cDia		:= "10"
	Local cCidade	:= "S�o Paulo"
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
				Aviso("Aten��o", "Alterne para o programa do Ms-Word para visualizar o contrato contrato" + cContrato + ".doc ou clique no botao para fechar.", {"Fechar"})
				OLE_SaveAsFile(hWord,cPath+"Contrato" + cContrato + ".doc")
			Endif
		
	EndMsOle()

	OLE_CloseLink( hWord )//fecha o Link com o Word

Return


Static Function AjustaSx1()
	PutSx1(cPerg,"01","Arquivo Modelo ","","","mv_ch1","C",99,0,0,"G","","DIR","","","mv_par01")
	PutSx1(cPerg,"02","Pasta Destino  ","","","mv_ch2","C",99,0,0,"G","","HSSDIR","","","mv_par02")
	PutSX1(cPerg,"03","Sa�da          ","","","mv_ch3","N",01,0,0,"C","","","","","mv_par03","Impressora", "", "", "","Arquivo")
Return
