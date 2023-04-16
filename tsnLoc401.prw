


//Bibliotecas
#Include "TOTVS.ch"
#Include "TopConn.ch"
 
/*/{Protheus.doc} User Function zTstDot
Função de teste, para mostrar como inserir linhas dinamicamente em um modelo dot
@type  Function
@author Atilio
@since 29/05/2021
/*/
 
User Function zTstDot(param_name)
    Local aArea := GetArea()
    Local cArquivo := "modelo.dotx"
    Local cArqSrv := "\x_dots\" + cArquivo
    Local cDirTmp := GetTempPath()
    Local nHandWord
    Local nItens := 0
    Local cQuery := ""
 
    //Copia o arquivo da protheus data para o pc do usuário
    CpyS2T(cArqSrv, cDirTmp)
 
    //Cria um ponteiro e já chama o arquivo
    nHandWord := OLE_CreateLink()
    OLE_NewFile(nHandWord, cDirTmp + cArquivo)
     
    //Setando o conteúdo das DocVariables
    OLE_SetDocumentVar(nHandWord, "DataHoje", dToC(Date()))
    OLE_SetDocumentVar(nHandWord, "HoraHoje", Time())
    OLE_SetDocumentVar(nHandWord, "Autor", UsrRetName(RetCodUsr()))
 
    //Busca 10 produtos que sejam do tipo PA e que não estejam bloqueados
    cQuery += " SELECT TOP 10 " + CRLF
    cQuery += "     B1_COD, " + CRLF
    cQuery += "     B1_DESC, " + CRLF
    cQuery += "     B1_UM " + CRLF
    cQuery += " FROM " + CRLF
    cQuery += "     " + RetSQLName("SB1") + " SB1 " + CRLF
    cQuery += " WHERE " + CRLF
    cQuery += "     B1_FILIAL = '" + FWxFilial('SB1') + "' " + CRLF
    cQuery += "     AND B1_MSBLQL != '1' " + CRLF
    cQuery += "     AND B1_TIPO = 'PA' " + CRLF
    cQuery += "     AND SB1.D_E_L_E_T_ = ' ' " + CRLF
    cQuery += " ORDER BY " + CRLF
    cQuery += "     B1_COD " + CRLF
    TCQuery cQuery New Alias "QRY_SB1"
 
    //Enquanto houver dados na query
    While ! QRY_SB1->(EOF())
        //Incrementa o total de itens
        nItens++
 
        //Define as variaveis das celulas 1, 2, 3 da linha atual no documento
        OLE_SetDocumentVar(nHandWord, 'ITEM1' + cValToChar(nItens), QRY_SB1->B1_COD )
        OLE_SetDocumentVar(nHandWord, 'ITEM2' + cValToChar(nItens), QRY_SB1->B1_DESC )
        OLE_SetDocumentVar(nHandWord, 'ITEM3' + cValToChar(nItens), QRY_SB1->B1_UM )
 
        QRY_SB1->(dbSkip())
    EndDo
    QRY_SB1->(DbCloseArea())
 
    //Define a quantidade de produtos e executa a macro para criar as linhas
    OLE_SetDocumentVar(nHandWord, 'QtdePro', cValToChar(nItens))
    OLE_ExecuteMacro(nHandWord, "tabItens")
 
    //Atualizando campos
    OLE_UpdateFields(nHandWord)
     
    //Monstrando um alerta
    MsgAlert('O arquivo gerado foi <b>Salvo</b>?<br>Ao clicar em OK o Microsoft Word será <b>fechado</b>!','Atenção')
     
    //Fechando o arquivo e o link
    OLE_CloseFile(nHandWord)
    OLE_CloseLink(nHandWord)
 
    RestArea(aArea)
Return
