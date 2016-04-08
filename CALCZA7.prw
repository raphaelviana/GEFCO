/*
Teste Raphael
/*


/*
Teste Raphael
/*


/*
Teste Raphael
/*


#INCLUDE "Totvs.ch"      
#INCLUDE "TOPConn.ch"                                                                                    

CLASS oCalcX
	DATA aCols 		as Array
	DATA aHeader 	as Array
	DATA nAT		as Integer
	METHOD NEW() CONSTRUCTOR
EndClass
Method New() class oCalcX
	::aCols := {}      
	::aHeader := {}      
Return Self

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³ CalcZA7  ºAutor  ³Itup                º Data ³  15/12/15   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Neste fonte contem as rotinas: fCalcZA7V e fCalcZA7C       º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ Itup / Gefco                                               º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function CalcZA7V()
Local aColsAux := {}
Local cOrig,cDest,cTpTab,cTpFlu,dDtIni,cViagem
Local cDevEmb,cAdic,cEmerg
Local cTpTran
Local cCodCli	:= 	AllTrim(ZA0->ZA0_CLI)
Local cLojCli	:=  AllTrim(ZA0->ZA0_LOJA)
Local cTabVen	//:=  AllTrim(ZA0->ZA0_TABVEN)
Local cVerTab
Local cTabCom
Local cVerTabC
Local nVALOR,nVEMERG,nPEMERG := 0
Local mSQL	:= ""
Local _VTotal := 0
Local _Ret := .t. 
Local _nX   		// Uso no calculo do formato 2
Local _nRat 		// Quantiade de rateios para o calculo do formato 2
Local _cCliTmp 		// Uso no calculo do formato 2
Local _cLojTmp 		// Uso no calculo do formato 2
Local _cCliDes 		// Uso no calculo do formato 3
Local _cLojDes 		// Uso no calculo do formato 3
Local _nPerRet		// Percentual do retorno para o formato 4
Local _nPerIda      // Percentual da ida para o formato 4
Local _nPerPec      // Percentual da peca para o formato 4
Local _nKm			// Quantidade de KM
Local _nVlrK		// Valor por KM
Local _nTotKm		// Valor total por KM
Local _nTotReg		// Contador de quantas linhas tem a ZTC
Local _aTemp

	cViagem   := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_OBRA"})]  //aColsAux[Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_OBRA"})] 
	cTpTrans  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TPTRAN"})] //aColsAux[Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TPTRAN"})] 
	cTpFluxo  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TIPFLU"})] //aColsAux[Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TIPFLU"})]
	dDtIni	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_DTINI" })] //aColsAux[Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_DTINI"})] 
	cEmerg	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_EMERG2"})] //aColsAux[Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_EMERG2"})] 
	cTabVen	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TABVEN"})] //aColsAux[Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TABVEN"})] 		
	cVerTab	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_VERVEN"})] //aColsAux[Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_VERVEN"})] 								
	cITTabV	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_ITTABV"})] //aColsAux[Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_VERVEN"})] 								

//	cTES	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TES"})]
	cCONPAG	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_CONPAG"})]
	cTPFRET	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TPFRET"})]
	cCCGEFC	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_CCGEFC"})]
	cCCCLIE	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_CCCLIE"})]
	cOI	  	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_OI" })]
	cCONTA 	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_CONTA"})]
	cTIPDES   := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TIPDES"})]                   
	cTpTran	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|Alltrim(x[2])=="ZA6_TRANSP"})] 
    
    _cObra	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_OBRA"})]
	_cSeq	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_SEQTRA"})]
                   

	// posicionar no SZ7 para encontrar o conteudo do ZA7_EMERG2 by Frank Z Fuga 17/11/2015
	For _nX:=1 to Len(oDlgCar:aCols)
		If !oDlgCar:aCols[_nX][len(oDlgCar:aHeader)+1]
	    	If oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_OBRA"})] == _cObra .and. oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQTRA"})] == _cSeq
	    		cEmerg := oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_EMERG2"})]
	    		Exit
	    	EndIF
		EndIF
	Next

/*
	if Len(oDlgCon:aCols) ==0 
		Alert("Na aba conjunto transportador não esta preenchida.")
		Return
	Else
		cTpTran := oDlgCon:aCols[1][Ascan(oDlgCon:aHeader,{|x|Alltrim(x[2])=="ZAE_TRANSP"})]
		if Empty(Alltrim(cTpTran))
			Alert("Na aba conjunto transportador não esta preenchida.")
			Return                                                                         
		Endif	 
	Endif
*/

	if Empty(Alltrim(cTpTran))
		Alert("O Campo Transportador não esta preenchida.")
		Return(.f.)
	Endif

/*
	if Empty(Alltrim(cTES))
		Alert("O Campo Tipo Saida não esta preenchida.")
		Return(.f.)
	Endif
*/

	if Empty(Alltrim(cCONPAG))
		Alert("O campo Condições de Pagamento não esta preenchida.")
		Return(.f.)
	Endif

	if Empty(Alltrim(cTPFRET))
		Alert("O campo Tipo de Frete não esta preenchida.")
		Return(.f.)
	Endif

	if Empty(Alltrim(cCCGEFC))
		Alert("O campo Centro de Custo GEFCO  não esta preenchida.")
		Return(.f.)
	Endif

	if Empty(Alltrim(cCCCLIE)) .and. SuperGetMv("IT_VALCLI",,"90136500") == cCodCli+cLojCli
		Alert("O campo Centro de Custo Cliente  não esta preenchida.")
		Return(.f.)
	Endif

	if Empty(Alltrim(cOI)) .and. SuperGetMv("IT_VALCLI",,"90136500") == cCodCli+cLojCli
		Alert("O campo Ordem Interna não esta preenchida.")
		Return(.F.)
	Endif

	if Empty(Alltrim(cCONTA)) .and. SuperGetMv("IT_VALCLI",,"90136500") == cCodCli+cLojCli
		Alert("O campo Conta PSA não esta preenchida.")
		Return(.F.)
	Endif

	if Empty(Alltrim(cTIPDES))
		Alert("O campo Tipo Despesa GEFCO não esta preenchida.")
		Return(.F.)
	Endif

/*
	DbSelectArea("ZAE")
	DbSetOrder(1)
	DbSeek(xFilial("ZAE")+ZA0->ZA0_PROJET+cVIAGEM)
	cTpTran := ZAE->ZAE_TRANSP
*/

	if Empty(cTabVen) .or. Empty(cVerTab) .or. Empty(cITTabV)
		Alert("Favor verifique a Tabela de Venda!")
		Return .F.
	Endif                               
	                                      
	
	// Posicionar na tabela para ver o tipo da regra para calculo
	// Frank - 12/11/15
	ZT0->( dbSetOrder(1) )	//  ZT0_FILIAL+ZT0_CODTAB+ZT0_VERTAB+ZT0_CODCLI+ZT0_LOJCLI+ZT0_TIPTAB+ZT0_ITEMTB
	ZT0->( dbSeek( xFilial("ZT0") + cTabVen + cVerTab + cCodCLi + cLojCli + cTpTrans + cITTabV , .T. ) )

	If ZT0->ZT0_TIPREG == "1"
		// Regra Renova 1            
		
		For nX := 1 To Len(oDlgCar:aCols)
			cDevEmb := oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})] //aColsAux[Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})] 
			cAdic   := oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_ADICIO"})] //aColsAux[Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_ADICIO"})]
			_Munic  := oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_MUNICI"})] //aColsAux[Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_MUNICI"})]
			
			cEmerg  := oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_EMERG2"})]        

			if Select("TZA2") > 0
				TZA2->(dbCloseArea())
			endif

			mSQL := "SELECT ZA2_CODIGO,ZA2_DESCRI,ZA2_ESTADO,ZA2_CDPAIS,ZA2_CODMUN, "
			mSQL += " ZT0_CODTAB,ZT0_VERTAB,ZT0_ITEMTB,ZT0_VALOR,ZT0_VEMERG,ZT0_TIPVEI, "
			mSQL += " ZT0_PEMERG,ZT0_VTXEMB,ZT0_PTXEMB,ZA2.R_E_C_N_O_ AS ZA2RECNO, "
			mSQL += " ZT0_VRTEMB,ZT0_PRTEMB,ZT0_VREEMB,ZT0_PREEMB "
			//mSQL += " FROM "+RetSQLName("ZA2")+" ZA2 LEFT JOIN "+RetSQLName("ZT0")+" ZT0 "
			mSQL += " FROM "+RetSQLName("ZA2")+" ZA2 INNER JOIN "+RetSQLName("ZT0")+" ZT0 "
			mSQL += " ON ZT0_CODORI=ZA2_CODIGO AND ZT0_CODTAB='"+cTabVen+"' AND ZT0_VERTAB='"+cVerTab+"'"
			mSQL += " AND ZT0_ITEMTB='"+cITTabV+"' "
			mSQL += " AND ZT0_CODCLI='"+cCodCLi+"' AND ZT0_LOJCLI='"+cLojCli+"' "
			mSQL += " AND ZT0_TIPTAB='"+cTpTrans+"' AND ZT0_TIPFLU='"+cTpFluxo+"' "
			mSQL += " WHERE ZT0_FILIAL='"+xFilial("ZT0")+"' AND ZT0.D_E_L_E_T_=' ' "
			mSQL += " AND ZA2_FILIAL='"+xFilial("ZA2")+"' AND ZA2.D_E_L_E_T_=' ' "
			mSQL += " AND ZT0_INIVIG <='"+DtoS(dDtIni)+"' AND ( ZT0_FIMVIG >='"+DtoS(dDtIni)+"'  OR ZT0_FIMVIG ='' ) "
			mSQL += " AND ZT0_TIPVEI='"+cTpTran+"' AND ZT0_MSBLQL='2' " //AND ZA2_CODIGO='"+_Munic+"' "
			mSQL += " GROUP BY  ZA2_CODIGO,ZA2_DESCRI,ZA2_ESTADO,ZA2_CDPAIS, "
			mSQL += " ZA2_CODMUN,ZA2.R_E_C_N_O_,ZT0_CODTAB,ZT0_VERTAB,ZT0_ITEMTB, "
			mSQL += " ZT0_VALOR,ZT0_VEMERG,ZT0_TIPVEI,ZT0_PEMERG,ZT0_VTXEMB,ZT0_PTXEMB, "
			mSQL += " ZT0_VRTEMB,ZT0_PRTEMB,ZT0_VREEMB,ZT0_PREEMB "
			dbUseArea( .T., "TOPCONN", TCGENQRY(,,mSQL), "TZA2", .F., .T. )

			if ! EOF()
				_vFrete := TZA2->ZT0_VALOR
				nVEMERG := TZA2->ZT0_VEMERG
				nPEMERG := TZA2->ZT0_PEMERG 
				_VTXEMB := TZA2->ZT0_VTXEMB
				_PTXEMB := TZA2->ZT0_PTXEMB
				_VRTEMB := TZA2->ZT0_VRTEMB
				_PRTEMB := TZA2->ZT0_PRTEMB
				_VREEMB := TZA2->ZT0_VREEMB
				_PREEMB := TZA2->ZT0_PREEMB
				_ItemTB := TZA2->ZT0_ITEMTB
 			else
				Loop    
			endif
			TZA2->(dbCloseArea())

			if cEmerg =="N" .AND. cDevEmb == "P" .AND. cAdic =="N" // ITEM 4.4
				nValor := _vFrete
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := nValor						
			Elseif cEmerg =="S" .AND. cDevEmb == "P" .AND. cAdic =="N"  // ITEM 4.5 
				if nVEMERG == 0 
					nValor :=_vFrete+(_vFrete * (nPEMERG/100))
				Else
					nValor := nVEMERG
				Endif
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := nValor						
			Elseif cEmerg =="N" .AND. cDevEmb == "I" .AND. cAdic =="N"
				if _VTXEMB == 0 
					nValor := _vFrete * (_PTXEMB/100)
				Else
					nValor := _VTXEMB
				Endif
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := nValor						
			Elseif cEmerg =="N" .AND. cDevEmb == "R" .AND. cAdic =="N" // ITEM 4.10
				if _VRTEMB == 0
					nValor := _vFrete+(_vFrete * (_PRTEMB/100))
				Else
					nValor := _VRTEMB
				Endif	
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := nValor						
			Elseif cEmerg =="S"  .AND. cDevEmb == "R" .AND. cAdic =="N"   // ITEM 4.11

				if _VREEMB == 0 
					nValor := _vFrete+(_vFrete * (_PREEMB/100))
				Else
					nValor := _VREEMB
				Endif
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := nValor						

			Elseif cEmerg =="N"  .AND. cDevEmb == "P" .AND. cAdic =="S"	 // ITEM 4.6				
				IF SELECT("TZTA") > 0
					dbSelectArea("TZTA")
					TZTA->(dbCloseArea())
				Endif
  

				mSQL := "SELECT ZTA_TABVEN,ZTA_VERTAB,ZTA_CODMUN,ZTA_VALOR,ZTA_PTXEMB,ZTA_VTXEMB,ZTA_PRTEMB,ZTA_VRTEMB "    
				mSQL += "FROM "+RetSQLName("ZTA")+" ZTA "
				mSQL += " WHERE ZTA_FILIAL ='"+xFilial("ZTA")+"' AND D_E_L_E_T_=' ' "
				mSQL += " AND ZTA_TABVEN='"+cTabVen+"' "
				mSQL += " AND ZTA_VERTAB='"+cVerTab+"' AND ZTA_CODMUN='"+_Munic+"'"
				mSQL += " AND ZTA_ITTABV='"+cITTabV+"' AND ZTA_MSBLQL='2' " 
				dbUseArea( .T., "TOPCONN", TCGENQRY(,,mSQL), "TZTA", .F., .T. )

				dbSelectArea("TZTA")
				TZTA->(dbGoTop())
				If TZTA->(!EoF())
					nValor := TZTA->ZTA_VALOR
					oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := nValor												
				Else
					oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := 0
					Loop
				Endif
				TZTA->(dbCloseArea())

			Elseif cEmerg =="S"  .AND. cDevEmb == "P" .AND. cAdic =="S"	 // ITEM 4.7
						
				IF SELECT("TZTA") > 0
					dbSelectArea("TZTA")
					TZTA->(dbCloseArea())
				Endif

				mSQL := "SELECT ZTA_TABVEN,ZTA_VERTAB,ZTA_CODMUN,ZTA_VALOR,ZTA_PTXEMB,ZTA_VTXEMB,ZTA_PRTEMB,ZTA_VRTEMB "    
				mSQL += " FROM "+RetSQLName("ZTA")+" ZTA "
				mSQL += " WHERE ZTA_FILIAL ='"+xFilial("ZTA")+"' AND D_E_L_E_T_=' ' "
				mSQL += " AND ZTA_TABVEN='"+cTabVen+"' "
				mSQL += " AND ZTA_VERTAB='"+cVerTab+"' AND ZTA_CODMUN='"+_Munic+"'"
				mSQL += " AND ZTA_ITTABV='"+cITTabV+"' AND ZTA_MSBLQL='2' " 						 
   				dbUseArea( .T., "TOPCONN", TCGENQRY(,,mSQL), "TZTA", .F., .T. )

				dbSelectArea("TZTA")
				TZTA->(dbGoTop())
				IF TZTA->(!EoF())
					nValor := TZTA->ZTA_VALOR
					oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := nValor						
				else
					oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := 0
					Loop	
				Endif

				TZTA->(dbCloseArea())

			Elseif cEmerg =="N"  .AND. cDevEmb == "I" .AND. cAdic =="S"	 // ITEM 4.9
						
				IF SELECT("TZTA") > 0
					dbSelectArea("TZTA")
					TZTA->(dbCloseArea())
				Endif

				mSQL := "SELECT ZTA_TABVEN,ZTA_VERTAB,ZTA_CODMUN,ZTA_VALOR,ZTA_PTXEMB,ZTA_VTXEMB,ZTA_PRTEMB,ZTA_VRTEMB "    
				mSQL += "FROM "+RetSQLName("ZTA")+" ZTA "
				mSQL += " WHERE ZTA_FILIAL ='"+xFilial("ZTA")+"' AND D_E_L_E_T_=' ' "
				mSQL += " AND ZTA_TABVEN='"+cTabVen+"' "
				mSQL += " AND ZTA_VERTAB='"+cVerTab+"' AND ZTA_CODMUN='"+_Munic+"'"
				mSQL += " AND ZTA_ITTABV='"+cITTabV+"'AND ZTA_MSBLQL='2' " 						 
				dbUseArea( .T., "TOPCONN", TCGENQRY(,,mSQL), "TZTA", .F., .T. )

				dbSelectArea("TZTA")
				TZTA->(dbGoTop())
				IF TZTA->(!EoF())
					_vFrete  := TZTA->ZTA_VALOR
					_VTXEMB  := TZTA->ZTA_VTXEMB
					_PTXEMB  := TZTA->ZTA_PTXEMB
					if _VTXEMB == 0 
						nValor := _vFrete * (_PTXEMB/100)
					Else
						nValor := _VTXEMB
					Endif
					oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := nValor
				else
					oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := 0
					Loop
				Endif

				TZTA->(dbCloseArea())
			Elseif cEmerg =="N"  .AND. cDevEmb == "R" .AND. cAdic =="S"	 // ITEM 4.12
				If Select("TZTA") > 0
					dbSelectArea("TZTA")
					TZTA->(dbCloseArea())
				Endif

				mSQL := "SELECT ZTA_TABVEN,ZTA_VERTAB,ZTA_CODMUN,ZTA_VALOR,ZTA_PTXEMB,ZTA_VTXEMB,ZTA_PRTEMB,ZTA_VRTEMB "    
				mSQL += "FROM "+RetSQLName("ZTA")+" ZTA "
				mSQL += " WHERE ZTA_FILIAL ='"+xFilial("ZTA")+"' AND D_E_L_E_T_=' ' "
				mSQL += " AND ZTA_TABVEN='"+cTabVen+"' "
				mSQL += " AND ZTA_VERTAB='"+cVerTab+"' AND ZTA_CODMUN='"+_Munic+"'"
				mSQL += " AND ZTA_ITTABV='"+cITTabV+"' AND ZTA_MSBLQL='2' " 
				dbUseArea( .T., "TOPCONN", TCGENQRY(,,mSQL), "TZTA", .F., .T. )

				dbSelectArea("TZTA")
				TZTA->(dbGoTop())
				If TZTA->(!EoF())
					nValor := TZTA->ZTA_VRTEMB
					If	nValor == 0	
						if _VRTEMB == 0
							nValor := _vFrete+(_vFrete * (_PRTEMB/100))
						Else
							nValor := _VRTEMB
						Endif											
					EndIf
					oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := nValor
				Else
					oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := 0
					Loop
				Endif
			
				TZTA->(dbCloseArea())				

			Elseif cEmerg =="S"  .AND. cDevEmb == "R" .AND. cAdic =="S"	 // ITEM 4.13

				if _VRTEMB == 0 
					nValor := _vFrete+(_vFrete * (_PRTEMB/100))
				Else
					nValor := _VRTEMB
				Endif
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := nValor						
			Endif
		Next nX    

	elseif ZT0->ZT0_TIPREG == "2"

		// Regra Renova 2
		// Frank - 12/11/15

		// Passo 1 - Encontrar a quantidade de clientes diferentes na pasta Coletas.
		// -------------------------------------------------------------------------
		If Len(oDlgCar:aCols) == 0
			Alert("Falta o preenchimento da pasta Coletas.")
			Return .F.
		EndIF                 
		// Buscar o primeiro codigo valido de clientes da pasta coletas
		For _nX:=1 to Len(oDlgCar:aCols)
			If !oDlgCar:aCols[01][len(oDlgCar:aHeader)+1]    
				_cCliTmp := oDlgCar:aCols[01][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CODCLI"})]
				_cLojTmp := oDlgCar:aCols[01][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJCLI"})] 		
				Exit
			EndIF
		Next
		// No caso de nao haver um cliente informado avisar para o usuario
		If empty(_cCliTmp)
			Alert("Falta o preenchimento da pasta Coletas - Identificação do Cliente.")
			Return .F.
		EndIF
		_nRat := 1
		For _nX:=1 to Len(oDlgCar:aCols)                                  
			If !oDlgCar:aCols[_nX][len(oDlgCar:aHeader)+1]    		
				If oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CODCLI"})]+oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJCLI"})] <> _cCliTmp+_cLojTmp
					_cCliTmp := oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CODCLI"})]
					_cLojTmp := oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJCLI"})] 		
					_nRat ++
				EndIF
			EndIF
		Next
		
		// Inicio do calculo
		For nX := 1 To Len(oDlgCar:aCols)
		
			If oDlgCar:aCols[nX][len(oDlgCar:aHeader)+1]    		
				Loop
			EndIF
			
			cDevEmb := oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})] //aColsAux[Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})] 
			cAdic   := oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_ADICIO"})] //aColsAux[Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_ADICIO"})]
			_Munic  := oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_MUNICI"})] //aColsAux[Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_MUNICI"})]

			cEmerg  := oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_EMERG2"})]                                 
            
			// Posicionar na tabela de vendas
			ZT0->( dbSetOrder(1) )	//  ZT0_FILIAL+ZT0_CODTAB+ZT0_VERTAB+ZT0_CODCLI+ZT0_LOJCLI+ZT0_TIPTAB+ZT0_ITEMTB
			ZT0->( dbSeek( xFilial("ZT0") + cTabVen + cVerTab + cCodCLi + cLojCli + cTpTrans + cITTabV , .T. ) )


			if ! ZT0->(EOF())
				_vFrete := ZT0->ZT0_VALOR / _nRat // Inserido a condicao do rateio
				
				nVEMERG := ZT0->ZT0_VEMERG / _nRat
				nPEMERG := ZT0->ZT0_PEMERG  
				_VTXEMB := ZT0->ZT0_VTXEMB
				_PTXEMB := ZT0->ZT0_PTXEMB
				_VRTEMB := ZT0->ZT0_VRTEMB
				_PRTEMB := ZT0->ZT0_PRTEMB
				_VREEMB := ZT0->ZT0_VREEMB
				_PREEMB := ZT0->ZT0_PREEMB
				_ItemTB := ZT0->ZT0_ITEMTB
 			else
				Loop    
			endif


			if cEmerg =="N" .AND. cDevEmb == "P" 
				nValor :=  _vFrete - ( (_vFrete*(_PRTEMB/100)) + (_vFrete*(_PTXEMB/100)) )
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := nValor	
			ElseIf cEmerg =="S" .AND. cDevEmb == "P" 
				If nPEMERG == 0
					nValor :=  nVEMERG - ( (nVEMERG*(_PRTEMB/100)) + (nVEMERG*(_PTXEMB/100)) )
				Else          
					nValor :=  _vFrete 
					nValor += nValor * (nPEMERG/100)                                            
					nValor -= ( (nValor*(_PRTEMB/100)) + (nValor*(_PTXEMB/100)) )
				EndIF                                                            
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := nValor	
			Elseif cEmerg =="N" .AND. cDevEmb == "I" 
				if _VTXEMB == 0 
					nValor := _vFrete * (_PTXEMB/100)
				Else
					nValor := _VTXEMB
				Endif
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := nValor						
			Elseif cEmerg =="S" .AND. cDevEmb == "I"                             
			
				If nPEMERG == 0
					nValor :=  nVEMERG 
				Else          
					nValor := _vFrete 
					nValor += nValor * (nPEMERG/100)                                            
				EndIF                                                            
			
				nValor := nValor * (_PTXEMB/100)
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := nValor						
	
	
			Elseif cEmerg =="N" .AND. cDevEmb == "R" 
				if _VRTEMB == 0
					nValor := _vFrete * (_PRTEMB/100)
				Else
					nValor := _VRTEMB
				Endif	
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := nValor						
			Elseif cEmerg =="S" .AND. cDevEmb == "R"                             
			
				If nPEMERG == 0
					nValor :=  nVEMERG 
				Else          
					nValor := _vFrete 
					nValor += nValor * (nPEMERG/100)                                            
				EndIF                                                            
			
				nValor := nValor * (_PRTEMB/100)
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := nValor						
			Endif

		Next nX    

	elseif ZT0->ZT0_TIPREG == "3"
		// Regra Renova 2
		// Cristiam - 15/12/15

		// Passo 1 - Encontrar a quantidade de clientes diferentes na pasta Coletas.
		// -------------------------------------------------------------------------
		If Len(oDlgCar:aCols) == 0
			Alert("Falta o preenchimento da pasta Coletas.")
			Return .F.
		EndIF                 
		// Buscar o primeiro codigo valido de clientes da pasta coletas
		For _nX:=1 to Len(oDlgCar:aCols)
			If !oDlgCar:aCols[01][len(oDlgCar:aHeader)+1]    
				_cCliTmp := oDlgCar:aCols[01][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CODCLI"})]
				_cLojTmp := oDlgCar:aCols[01][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJCLI"})]
				_cCliDes := oDlgCar:aCols[01][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CLIDES"})]
				_cLojDes := oDlgCar:aCols[01][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJDES"})]
				Exit
			EndIF
		Next
		// No caso de nao haver um cliente informado avisar para o usuario
		If empty(_cCliTmp) .or. empty(_cCliDes)
			Alert("Falta o preenchimento da pasta Coletas - Identificação do Cliente.")
			Return .F.
		EndIF
		_nRat := 1
		For _nX:=1 to Len(oDlgCar:aCols)                                  
			If !oDlgCar:aCols[_nX][len(oDlgCar:aHeader)+1]    		
				If	oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CODCLI"})] +;
					oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJCLI"})] +;
					oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CLIDES"})] +;
					oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJDES"})] != ;
					_cCliTmp + _cLojTmp + _cCliDes + _cLojDes

					_cCliTmp := oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CODCLI"})]
					_cLojTmp := oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJCLI"})]
					_cCliDes := oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CLIDES"})]
					_cLojDes := oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJDES"})]

					_nRat ++
				EndIF
			EndIF
		Next
		
		// Inicio do calculo
		For nX := 1 To Len(oDlgCar:aCols)
		
			If oDlgCar:aCols[nX][len(oDlgCar:aHeader)+1]    		
				Loop
			EndIF
			
			cDevEmb := oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})] //aColsAux[Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})] 
			cAdic   := oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_ADICIO"})] //aColsAux[Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_ADICIO"})]
			_Munic  := oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_MUNICI"})] //aColsAux[Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_MUNICI"})]

			cEmerg  := oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_EMERG2"})]                                 
            
			// Posicionar na tabela de vendas
			ZT0->( dbSetOrder(1) )	//  ZT0_FILIAL+ZT0_CODTAB+ZT0_VERTAB+ZT0_CODCLI+ZT0_LOJCLI+ZT0_TIPTAB+ZT0_ITEMTB
			ZT0->( dbSeek( xFilial("ZT0") + cTabVen + cVerTab + cCodCLi + cLojCli + cTpTrans + cITTabV , .T. ) )


			if ! ZT0->(EOF())
				_vFrete := ZT0->ZT0_VALOR / _nRat // Inserido a condicao do rateio
				
				nVEMERG := ZT0->ZT0_VEMERG / _nRat
				nPEMERG := ZT0->ZT0_PEMERG  
				_VTXEMB := ZT0->ZT0_VTXEMB
				_PTXEMB := ZT0->ZT0_PTXEMB
				_VRTEMB := ZT0->ZT0_VRTEMB
				_PRTEMB := ZT0->ZT0_PRTEMB
				_VREEMB := ZT0->ZT0_VREEMB
				_PREEMB := ZT0->ZT0_PREEMB
				_ItemTB := ZT0->ZT0_ITEMTB
 			else
				Loop    
			endif


			if cEmerg =="N" .AND. cDevEmb == "P" 
				nValor :=  _vFrete - ( (_vFrete*(_PRTEMB/100)) + (_vFrete*(_PTXEMB/100)) )
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := nValor	
			ElseIf cEmerg =="S" .AND. cDevEmb == "P" 
				If nPEMERG == 0
					nValor :=  nVEMERG - ( (nVEMERG*(_PRTEMB/100)) + (nVEMERG*(_PTXEMB/100)) )
				Else          
					nValor :=  _vFrete 
					nValor += nValor * (nPEMERG/100)                                            
					nValor -= ( (nValor*(_PRTEMB/100)) + (nValor*(_PTXEMB/100)) )
				EndIF                                                            
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := nValor	
			Elseif cEmerg =="N" .AND. cDevEmb == "I" 
				if _VTXEMB == 0 
					nValor := _vFrete * (_PTXEMB/100)
				Else
					nValor := _VTXEMB
				Endif
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := nValor						
			Elseif cEmerg =="S" .AND. cDevEmb == "I"                             
			
				If nPEMERG == 0
					nValor :=  nVEMERG 
				Else          
					nValor := _vFrete 
					nValor += nValor * (nPEMERG/100)                                            
				EndIF                                                            
			
				nValor := nValor * (_PTXEMB/100)
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := nValor						
	
	
			Elseif cEmerg =="N" .AND. cDevEmb == "R" 
				if _VRTEMB == 0
					nValor := _vFrete * (_PRTEMB/100)
				Else
					nValor := _VRTEMB
				Endif	
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := nValor						
			Elseif cEmerg =="S" .AND. cDevEmb == "R"                             
			
				If nPEMERG == 0
					nValor :=  nVEMERG 
				Else          
					nValor := _vFrete 
					nValor += nValor * (nPEMERG/100)                                            
				EndIF                                                            
			
				nValor := nValor * (_PRTEMB/100)
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := nValor						
			Endif

		Next nX                  

	elseif ZT0->ZT0_TIPREG == "4"
	
		// Frank Zwarg Fuga - 26/01/2016
		
		// Passo 1 - Calcular o ZA7_VALOR para todas as Pecas
		// --------------------------------------------------
		For _nX:=1 to Len(oDlgCar:aCols)	
			If !oDlgCar:aCols[_nX][len(oDlgCar:aHeader)+1]    		
				If oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})]=="P"
                    
					// PASSO 1 - Encontrar o ZA7_VALOR para a PECA posicionada
					// -------------------------------------------------------
					_nKm := U_KMMOD2("1", _nX, "1")      //oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SKMVEN"})]
					
					dbSelectArea("ZTC")
					DbOrderNickName("ITUPZTC001")
					_nVlrK := 0
					dbSeek(xFilial("ZTC")+ZT0->ZT0_CODTAB+ZT0->ZT0_VERTAB+ZT0->ZT0_ITEMTB)
					_aFaixa := {}      
					_lVRFIXO := .F.
					While !Eof() .and. ZTC_FILIAL+ZTC_TABVEN+ZTC_VERVEN+ZTC_ITTABV == xFilial("ZTC")+ZT0->ZT0_CODTAB+ZT0->ZT0_VERTAB+ZT0->ZT0_ITEMTB
						aadd(_aFaixa, {ZTC_FAIXAD, ZTC_FAIXAA, ZTC->ZTC_VRKMON, ZTC_VRFIXO, ZTC_VREONE })
						If ZTC_VRFIXO > 0
							_lVRFIXO := .T.
						EndIF
					    dbSkip()
					EndDo                                                                 
					              
					_nVlrK := 0
					_lSai := .F.
					For _nZ:=1 to len(_aFaixa)
						If _aFaixa[_nZ][01]==0 .and. _aFaixa[_nZ][02]==0
							_nVlrK := _aFaixa[_nZ][03] * _nKm // VRKMON
							Exit  
						Else                         
						
							_lSai := .F.
							If !_lVRFIXO                            
								If _aFaixa[_nZ][01] <= _nKm .and. _aFaixa[_nZ][2] >= _nKm
									_nVlrK := _aFaixa[_nZ][03] * _nKm // VRKMON
									Exit
								EndIF
							Else  
								If _aFaixa[_nZ][01] <= _nKm .and. _aFaixa[_nZ][2] >= _nKm
//									If _nZ == len(_aFaixa) 
									If _nZ == len(_aFaixa) .and. _nZ != 1			// evitar o erro de array out of bounds - Cristiam Rossi em 28/02/2016
										_nTemp  := _nKm - _aFaixa[_nZ-1][2]
										_nTemp1 := _nTemp * _aFaixa[_nZ][05]
										_nTemp2 := _nKm - _nTemp
										For _nT:=1 to len(_aFaixa)
											If _aFaixa[_nT][01] <= _nTemp2 .and. _aFaixa[_nT][2] >= _nTemp2
												_nVlrK := _nTemp1 + _aFaixa[_nT][04]
												_lSai := .T.
												Exit
											EndIF
										Next
										If _lSai
											Exit
										EndIF
									Else
										_nVlrK := _aFaixa[_nZ][04] // VRFIXO e dentro de uma faixa que não é a ultima
										_lSai := .T.
										Exit
									EndIf
								EndIF						
							EndIf
							If _lSai
								Exit
							EndIF

						EndIf
							
					Next
					
					oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := _nVlrK
					//_aRet := U_KMMOD2("1", _nX, "2")
					
					//If _aRet[1] == "S"
						// Se for calculo por rateio, vamos processar somente para a primeira linha do tipo PECA,
						// pois o valor ja esta aglutinado.
						
						// Zerar das demais linhas para nao conflitar no calculo do rateio.
					//	For _nZ:=1 to Len(oDlgCar:aCols)
					//		If _nZ <> _nX
					//			oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := 0
					//		EndIF
					//	Next
					//	Exit
					//EndIF
					 
			    EndIF
			EndIF
		Next				 
		
		// PASSO 2 - Calcular o ZA7_VALOR para os retornos.
		// ------------------------------------------------
		For _nX:=1 to Len(oDlgCar:aCols)	
			If !oDlgCar:aCols[_nX][len(oDlgCar:aHeader)+1] 
				If oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})]=="R"		
					_nVlrK := 0   		
					
					// Encontrar a peca correspondente ao registro de retorno.
					// -------------------------------------------------------
					_cCliTmp := oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CODCLI"})]
					_cLojTmp := oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJCLI"})]
					_cCliDes := oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CLIDES"})]
					_cLojDes := oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJDES"})]
					_nKmPeca := 0
					_nVlPeca := 0
					For _nY:=1 to Len(oDlgCar:aCols)	
						If !oDlgCar:aCols[_nY][len(oDlgCar:aHeader)+1] 
							If oDlgCar:aCols[_nY][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})]=="P"		
								If oDlgCar:aCols[_nY][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CODCLI"})]==_cCliTmp
									If oDlgCar:aCols[_nY][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJCLI"})]==_cLojTmp
										If oDlgCar:aCols[_nY][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CLIDES"})]==_cCliDes
											If oDlgCar:aCols[_nY][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJDES"})]==_cLojDes
												_nKmPeca := oDlgCar:aCols[_nY][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SKMVEN"})]
												_nVlPeca := oDlgCar:aCols[_nY][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})]
												Exit
											EndIf
										EndIF
									EndIF
								EndIF
							EndIF
						EndIF
					Next       
					
					// Encontrar as faixas validas.
					// ----------------------------
					dbSelectArea("ZTC")
					DbOrderNickName("ITUPZTC001")
					_nVlrK := 0
					dbSeek(xFilial("ZTC")+ZT0->ZT0_CODTAB+ZT0->ZT0_VERTAB+ZT0->ZT0_ITEMTB)
					_aFaixa := {}      
					_lVRFIXO := .F.
					While !Eof() .and. ZTC_FILIAL+ZTC_TABVEN+ZTC_VERVEN+ZTC_ITTABV == xFilial("ZTC")+ZT0->ZT0_CODTAB+ZT0->ZT0_VERTAB+ZT0->ZT0_ITEMTB
						aadd(_aFaixa, {ZTC_FAIXAD, ZTC_FAIXAA, ZTC->ZTC_VRKMON, ZTC_VRFIXO, ZTC_VREONE, ZTC_VRKMRO, ZTC_PEKMRO, ZTC_VRFIXR, ZTC_PERR, ZTC_VREROU })
					    dbSkip()
					EndDo

					_nTemp := 0					
					For _nZ:=1 to len(_aFaixa)
						_lSai := .F.
						// Calculo por faixa = 0
						// -----------------------------------------
						If _aFaixa[_nZ][01]==0 .and. _aFaixa[_nZ][02]==0
							If _aFaixa[_nZ][06] > 0 // VRKMRO
								_nTemp := (_nKmPeca * _aFaixa[_nZ][06]) - _nVlPeca
							Else
								_nTemp := (_nKmPeca * ((  (_aFaixa[_nZ][03]*_aFaixa[_nZ][07])/100   )+_aFaixa[_nZ][03])) - _nVlPeca
							EndIF

							_nKm := U_KMMOD2("1", _nX, "1") //oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SKMVEN"})]
							
							If _aFaixa[_nZ][06] > 0 // VRKMRO
								_nTemp := (_nKm * _aFaixa[_nZ][06]) + _nTemp
							Else
								_nTemp := (_nKm * (( (_aFaixa[_nZ][03]*_aFaixa[_nZ][07])/100  )+_aFaixa[_nZ][03])) + _nTemp
							EndIF

							Exit  
						EndIF  
						// Calculo por faixa preenchida
						// ----------------------------
						// CASO 1 - o valor do KM se enquadra em uma faixa sem ser a ultima
						// ----------------------------------------------------------------
						_nKm := U_KMMOD2("1", _nX, "1") //oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SKMVEN"})]
						
						If _aFaixa[_nZ][01] <= _nKm .and. _aFaixa[_nZ][2] >= _nKm .and. _nZ < len(_aFaixa)
							If _aFaixa[_nZ][04] > 0 // Valor fixo
								If _aFaixa[_nZ][08] > 0 // VRFIXR
									_nTemp := _aFaixa[_nZ][08] - _aFaixa[_nZ][04]
									Exit
								Else
									_nTemp := (_aFaixa[_nZ][04] * _aFaixa[_nZ][09])/100
									Exit
								EndIF
							Else     
								If _aFaixa[_nZ][06] == 0 // VRKMRO
									_nTemp1 := ((_aFaixa[_nZ][03] * _aFaixa[_nZ][07])/100)+_aFaixa[_nZ][03]
									_nTemp2 := (_nKmPeca * _nTemp1) - _nVlPeca
									_nTemp  := (_nTemp1 * _nKm) + _nTemp2
									Exit
								Else                                     
									_nTemp1 := _aFaixa[_nZ][06] * _nKmPeca
									_nTemp2 := _nTemp1 - _nVlPeca
									_nTemp  := (_nKm * _aFaixa[_nZ][06]) + _nTemp2
									Exit
								EndIF 
							EndIF
						EndIF
						// CASO 2 - o valor do KM se enquadra na ultima faixa 
						// ----------------------------------------------------------------
						If _aFaixa[_nZ][01] <= _nKm .and. _aFaixa[_nZ][2] >= _nKm .and. _nZ == len(_aFaixa)
							If _aFaixa[_nZ-1][04] > 0 // vlr fixo
								// posicionar na faixa anterior
								If _aFaixa[_nZ-1][08] > 0 // VRFIXR
									_nTemp1 := _aFaixa[_nZ-1][08] - _aFaixa[_nZ-1][04]
									_nTemp2 := _nKmPeca - _aFaixa[len(_aFaixa)-1][02]
									_nTemp3 := _nTemp2 * _aFaixa[_nZ][05] // VREONE
									_nTemp4 := _nTemp2 * _aFaixa[_nZ][10] // VREROU
									_nTemp5 := _nTemp4 - _nTemp3
									_nTemp6 := _nKm - _aFaixa[len(_aFaixa)-1][02]
									_nTemp  := _nTemp6 * _aFaixa[_nZ][10] 
									_nTemp  := _nTemp + _nTemp5 + _nTemp1
									exit
								Else    
									_nTemp1 := (_aFaixa[_nZ-1][04]*_aFaixa[_nZ-1][09])/100
									_nTemp2 := _nKmPeca - _aFaixa[len(_aFaixa)-1][02]
									_nTemp3 := _nTemp2 * _aFaixa[_nZ][05] // VREONE
									_nTemp4 := _nTemp2 * _aFaixa[_nZ][10] // VREROU
									_nTemp5 := _nTemp4 - _nTemp3
									_nTemp6 := _nKm - _aFaixa[len(_aFaixa)-1][02]
									_nTemp  := _nTemp6 * _aFaixa[_nZ][10] 
									_nTemp  := _nTemp + _nTemp5 + _nTemp1
									exit
								EndIF
								
							Else                               
								
								_nTemp1 := _nKmPeca * _aFaixa[_nZ][06]  // ZTC_VRKMRO
								_nTemp2 := _nTemp1 - _nVlPeca
								_nTemp3 := _nKm * _aFaixa[_nZ][06]
								_nTemp  := _nTemp3 + _nTemp2
								Exit
								
							EndIF
						EndIF
						
					Next
					
					oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := _nTemp
					//_aRet := U_KMMOD2("1", _nX, "2")
					//If _aRet[2] == "S"
						// Se for calculo por rateio, vamos processar somente para a primeira linha do tipo RETORNO,
						// pois o valor ja esta aglutinado.
						
						// Zerar das demais linhas para nao conflitar no calculo do rateio.
					//	For _nZ:=1 to Len(oDlgCar:aCols)
					//		If _nZ <> _nX .and. oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})]=="R"		 
					//			oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := 0
					//		EndIF
					//	Next
					//	Exit
					//EndIF
					
				EndIf 
			EndIF
		Next

		//------------ Considerar Cálculo Emergencial
		nPEMERG := 0
		nPERFER := 0
		// Posicionar na tabela de vendas
		ZT0->( dbSetOrder(1) )	//  ZT0_FILIAL+ZT0_CODTAB+ZT0_VERTAB+ZT0_CODCLI+ZT0_LOJCLI+ZT0_TIPTAB+ZT0_ITEMTB
		if ZT0->( dbSeek( xFilial("ZT0") + cTabVen + cVerTab + cCodCLi + cLojCli + cTpTrans + cITTabV , .T. ) )
			nPEMERG := ZT0->ZT0_PEMERG
			nPERFER := ZT0->ZT0_PERFER
		endif

		if gdFieldGet("ZA6_FERIAV", oDlgTra:nAt, .F., oDlgTra:aHeader, oDlgTra:aCols ) != "S"	// não vai tratar o % Feriado
			nPERFER := 0
		endif

		For _nX := 1 to Len(oDlgCar:aCols)
			if ! gdDeleted( _nX, oDlgCar:aHeader, oDlgCar:aCols)
				if gdFieldGet("ZA7_EMERG2", _nX, .F., oDlgCar:aHeader, oDlgCar:aCols ) == "S" .and. nPEMERG > 0	// se for SIM, adicionar % de Emergência
					nValor := gdFieldGet("ZA7_VALOR", _nX, .F., oDlgCar:aHeader, oDlgCar:aCols ) 
					nValor += nValor * nPEMERG / 100
			    	gdFieldPut("ZA7_VALOR", nValor, _nX, oDlgCar:aHeader, oDlgCar:aCols)
				endif

				if nPERFER > 0	// se for SIM, adicionar % de Feriado
					nValor := gdFieldGet("ZA7_VALOR", _nX, .F., oDlgCar:aHeader, oDlgCar:aCols ) 
					nValor += nValor * nPERFER / 100
			    	gdFieldPut("ZA7_VALOR", nValor, _nX, oDlgCar:aHeader, oDlgCar:aCols)
				endif

			endif
		next
		
		_aRet := U_KMMOD2("1", _nX, "2")
		If _aRet[1] == "S" // Tratamento especial para rateio por KM/PESO
		
			// Tratamento especial para rateio por KM ou PESO
			// ----------------------------------------------------------------------------------------------------------------

			// Detalhamento do Retorno da função KMMOD2
			// cVenda , cCompra , cRatVen , cRatCom , nPerV , nPerC
			// cVenda  == "S" representa que havera rateio para o calculo da venda
			// cCompra == "S" representa que havera rateio para o calculo da compra
			// cRatVen == "N" representa que o rateio da venda sera por KM
			// cRatVen == "S" representa que o rateio da venda sera por PESO
			// cRatCom == "N" representa que o rateio da compra sera por KM
			// cRatCom == "S" representa que o rateio da compra sera por PESO
			// nPerV          representa o % do rateio da embalagem para venda
			// nPerC	      representa o % do rateio da embalagem para compra                  
	
			// PASSO 1 - Validacoes especiais para realizar o calculo
			// ----------------------------------------------------------------------------------------------------------------
			_lErro := .F.
			For _nX:=1 to Len(oDlgCar:aCols)    
				If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
					_aRet := U_KMMOD2("1", _nX, "2")
					If _aRet[1] == "S" // todos os rateios devem ter o seqcol preenchido
						If empty(oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})])
							_lErro := .T.
						EndIF
					EndIF
				EndIF
			Next
			If _lErro
				MsgStop("As sequencias da coleta nao foram preenchidas.","Atenção!")
				Return _Ret
			EndIf
			// ----------------------------------------------------------------------------------------------------------------
			// PASSO 2 - Encontrar o total de KM de todas as coletas associadas que tenha o campo SEQCOL preenchido
			_nTotKM := 0

			// Array com as faixas para recalculo do valor total do frete
			// ----------------------------------------------------------
			dbSelectArea("ZTC")
			DbOrderNickName("ITUPZTC001")
			dbSeek(xFilial("ZTC")+ZT0->ZT0_CODTAB+ZT0->ZT0_VERTAB+ZT0->ZT0_ITEMTB)
			_aFaixa := {}      
			While !Eof() .and. ZTC_FILIAL+ZTC_TABVEN+ZTC_VERVEN+ZTC_ITTABV == xFilial("ZTC")+ZT0->ZT0_CODTAB+ZT0->ZT0_VERTAB+ZT0->ZT0_ITEMTB
				aadd(_aFaixa, {ZTC_FAIXAD, ZTC_FAIXAA, ZTC->ZTC_VRKMRO, ZTC_VRFIXR })
			    dbSkip()                                           
			EndDo                                                  
			// ----------------------------------------------------               

			For _nX:=1 to Len(oDlgCar:aCols)    
				If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
					_aRet := U_KMMOD2("1", _nX, "2")
					If _aRet[1] == "S" 
						If !empty(oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})])
							_nTotKm  += oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SKMVEN"})]
						EndIF
					EndIf
				EndIF
			Next                     

			// Calculo do valor total
			_nTotVlr := 0
			If _nTotKm > 0
				_nTotVlr  := 0
				For _nZ:=1 to len(_aFaixa)
					If _aFaixa[_nZ][01]==0 .and. _aFaixa[_nZ][02]==0
						_nTotVlr  := _aFaixa[_nZ][03] * _nTotKm // VRKMON
						Exit  
					EndIf                         
					If _aFaixa[_nZ][01] <= _nTotKm .and. _aFaixa[_nZ][2] >= _nTotKm
						If _aFaixa[_nZ][04] > 0
							_nTotVlr  := _aFaixa[_nZ][04] // VRFIXO e dentro de uma faixa que não é a ultima
						Else
							_nTotVlr  := _aFaixa[_nZ][03] * _nTotKm
						EndIF
						Exit
					EndIF						
				Next
			EndIF                     
		
			// Encontrar quantos clientes diferentes existem na ZA7 para rateio quando nao tiver retorno ou peca associada.
			_aClientes := {}
			For _nX:=1 to Len(oDlgCar:aCols)    
				If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
					_aRet := U_KMMOD2("1", _nX, "2")
					If _aRet[1] == "S"              
						If !empty(oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})])
							If Len(_aClientes) == 0
								aadd(_aClientes,{oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CODCLI"})]+;
								              oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJCLI"})] })
							Else 
	
	        					_cTemp  := oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CODCLI"})] + oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJCLI"})]
	        					_lAchou := If( aScan(_aClientes, {|x| x[1] == _cTemp }) > 0, .T., .F.)

								If !_lAchou
									aadd(_aClientes,{oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CODCLI"})]+;
								              oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJCLI"})] })						
							    EndIF
							EndIF
						EndIF					
					EndIf
				EndIF
			Next
			_nRateio := 0
			If len(_aClientes) > 0
				_nRateio := _nTotVlr / len(_aClientes)
			EndIF
		
			// PASSO 2.1 - Encontrar o total de PESO de todas as coletas associadas que tenha o campo SEQCOL preenchido
			_nTotP := 0
			For _nX:=1 to Len(oDlgCar:aCols)    
				If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
					_aRet := U_KMMOD2("1", _nX, "2")
					If _aRet[1] == "S"              
						// Totalizar o peso somente para as pecas
						If !empty(oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})]) //.and. oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})] == "P"
							_nTotP += oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_PESO"})]
						EndIF
					EndIf
				EndIF
			Next
			// -------------------------------------------------------------------------------------------------------------
			// PASSO 3 - Geracao do temporario para realizacao dos calculos
			_aTemp := {}
			For _nX:=1 to Len(oDlgCar:aCols)    
				If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
					_aRet := U_KMMOD2("1", _nX, "2")
					If _aRet[1] == "S"
						If !empty(oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})])
							aadd(_aTemp, {oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})],;
							 			  oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})],;
							 			  oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CODCLI"})],;
							 			  oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJCLI"})],;   
							 			  oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CLIDES"})],;       
							 			  oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJDES"})],;
							 			  oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SKMVEN"})],;
							 			  0,;
							 			  " ",;
							 			  _aRet[6],;
							 			  oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_PESO"})] })
						EndIF
					EndIf
				EndIF
			Next
			_aTemp := aSort(_aTemp,,,{|x,y| x[1] < y[1] })

			// Verificar se existe algum item que nao tenha retorno ou peca.
			// Se todos tiverem, aplicar o rateio por KM ou PESO
			//_lPorKMPeso := .T.
			_lPorKMPeso := .F. // Frank em 02/03/16
			For _nX:=1 to Len(_aTemp)
				If _aTemp[_nX][02] == "R"
					_lTemProximo := .F.
					For _nY:=1 to Len(_aTemp)
						If _nY <> _nX .and. _aTemp[_nY][03]==_aTemp[_nX][03] .and. _aTemp[_nY][04]==_aTemp[_nX][04] .and. _aTemp[_nY][05]==_aTemp[_nX][05] .and. _aTemp[_nY][06]==_aTemp[_nX][06] .and. _aTemp[_nY][02]<>_aTemp[_nX][02]
							_lTemProximo 	:= .T.   
							Exit
						EndIF
					Next          
					If !_lTemProximo
						_lPorKMPeso := .F.
					EndIF
				EndIF
			Next                                   
		
			If _lPorKMPeso // Rateio por KM ou PESO

				// Calcular primeiro todos as coletas que nao tem rateio pela tabela de vendas
				For _nX:=1 to Len(_aTemp)
					_lTemProximo := .F.
					_nTemp       := 0
					If _aTemp[_nX][09] == "X"
						Loop
					EndIF
					For _nY:=1 to Len(_aTemp)
						If _nY <> _nX .and. _aTemp[_nY][03]==_aTemp[_nX][03] .and. _aTemp[_nY][04]==_aTemp[_nX][04] .and. _aTemp[_nY][05]==_aTemp[_nX][05] .and. _aTemp[_nY][06]==_aTemp[_nX][06] .and. _aTemp[_nY][02]<>_aTemp[_nX][02]
							_lTemProximo 	:= .T.   
							Exit
						EndIF
					Next
					If !_lTemProximo .and. _aTemp[_nX][2] == "P"           
						// Se nao tem proximo e for peca nao ha rateio, porem se for Retorno há o rateio por pecas por cliente diferente
						_aRet := U_KMMOD2("1", _nX, "2")
						If _aRet[3] == "N" // rateio por km
							_aTemp[_nX][08] := (_aTemp[_nX][07]/_nTotKM) *  _nTotVlr 
						Else // rateio por peso
							_aTemp[_nX][08] := (_aTemp[_nX][11]/_nTotP) *  _nTotVlr 
						EndIF
						// Atualizar a ZA7
						For _nZ:=1 to Len(oDlgCar:aCols)
							If !oDlgCar:aCols[_nZ][Len(oDlgCar:aHeader)+1]
								If oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})] == _aTemp[_nX][01]
									oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := _aTemp[_nX][08]
								EndIF
							EndIf
						Next
						_aTemp[_nX][09] := "X" 
					EndIF
				Next

				// ---------------------------------------------------------------------------------------------------------------
				// Calcular os que possuem rateio
				For _nX:=1 to Len(_aTemp)
					If _aTemp[_nX][09] == "X" // ja processado
						Loop
					EndIf
	
					// Encontrar o rateio de todas as coletas associadas
					// -----------------------------------------------------------------------------------------------------------
					_lTemProximo := .F.
					_nTemp       := 0
					_aRet := U_KMMOD2("1", _nX, "2")
					If _aRet[3] == "N" // rateio por km
						_nTemp4      := _aTemp[_nX][07]
					Else // rateio por peso
						_nTemp4      := _aTemp[_nX][11]			
					EndIf			
					// acumular em _nTemp4 o total dos rateios
					_aTemp[_nX][09] := "X"              
					For _nY:=1 to Len(_aTemp)
						If _nY <> _nX .and. _aTemp[_nY][03]==_aTemp[_nX][03] .and. _aTemp[_nY][04]==_aTemp[_nX][04] .and. _aTemp[_nY][05]==_aTemp[_nX][05] .and. _aTemp[_nY][06]==_aTemp[_nX][06] .and. _aTemp[_nY][02]<>_aTemp[_nX][02] .and. _aTemp[_nY][09] == " "
							_aRet := U_KMMOD2("1", _nY, "2")
							If _aRet[3] == "N" // rateio por km
								_nTemp4 		+= _aTemp[_nY][07]
							Else // rateio por peso
								_nTemp4 		+= _aTemp[_nY][11]
							EndIf
							_aTemp[_nY][09] := "X"
							Exit
						EndIF
					Next
					// -----------------------------------------------------------------------------------------------------------
					// Encontrar o rateio por igual para todas as coletas
					_aRet := U_KMMOD2("1", _nY, "2")
					If _aRet[3] == "N" // rateio por km
						_nTemp := (_nTemp4/_nTotKM) *  _nTotVlr 
					Else // rateio por peso
						_nTemp := (_nTemp4/_nTotP) *  _nTotVlr 			
					EndIF                                               
					// -----------------------------------------------------------------------------------------------------------
					// CASO 1 - estamos no retorno
					// Aplicar o percentual da ZT0 no retorno                                   
					_nTemp5 := 0
					If _aTemp[_nX][02] == "R"
						_nTemp1 := (_nTemp * _aTemp[_nX][10])/100 // % da tabela de venda para retorno
						If _aTemp[_nX][10] == 0
							MsgStop("O calculo foi executado com um erro, pois o % de retorno da tabela de vendas esta zerado.","Atenção!")
						EndIf           
						_nTemp5 := _aTemp[_nX][10] // armazenar o % aplicado para realizar a diferenca na peca
						// Atualizar a ZT7                 
						For _nZ:=1 to Len(oDlgCar:aCols)
							If !oDlgCar:aCols[_nZ][Len(oDlgCar:aHeader)+1]
								If oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})] == _aTemp[_nX][01]
									oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := _nTemp1
								EndIF
							EndIf
						Next
						// Procurar pela peca e aplicar o percentual da diferencao da ZT0
						For _nY:=1 to Len(_aTemp)
							If _nY <> _nX .and. _aTemp[_nY][03]==_aTemp[_nX][03] .and. _aTemp[_nY][04]==_aTemp[_nX][04] .and. _aTemp[_nY][05]==_aTemp[_nX][05] .and. _aTemp[_nY][06]==_aTemp[_nX][06] .and. _aTemp[_nY][02]=="P"
								_nTemp1	:= (_nTemp * (100-_nTemp5))/100 // diferenca do percentual da tabela de vendas
								_aTemp[_nY][09] := "X"
								// Atualizar a ZT7
								For _nZ:=1 to Len(oDlgCar:aCols)
									If !oDlgCar:aCols[_nZ][Len(oDlgCar:aHeader)+1]
										If oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})] == _aTemp[_nY][01]
											oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := _nTemp1
											Exit
										EndIF
									EndIf
								Next
								Exit
							EndIF
						Next     
					Else // é uma peca      
						// ---------------------------------------------------------------------------------------------------------
						// procurar pelo retorno
						_nTemp5 := 0
						For _nY:=1 to Len(_aTemp)
							If _nY <> _nX .and. _aTemp[_nY][03]==_aTemp[_nX][03] .and. _aTemp[_nY][04]==_aTemp[_nX][04] .and. _aTemp[_nY][05]==_aTemp[_nX][05] .and. _aTemp[_nY][06]==_aTemp[_nX][06] .and. _aTemp[_nY][02]=="R"
								_nTemp1 := (_nTemp * _aTemp[_nY][10])/100 // % da tabela de venda para retorno
								_nTemp5 := _aTemp[_nY][10]
								_aTemp[_nY][10] := "X"
								// Atualizar a ZT7                 
								For _nZ:=1 to Len(oDlgCar:aCols)
									If !oDlgCar:aCols[_nZ][Len(oDlgCar:aHeader)+1]
										If oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})] == _aTemp[_nY][01]
											oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := _nTemp1
											Exit
										EndIF
									EndIf
								Next
								Exit
							EndIF  
						Next
						// processar a peca           
						_nTemp1	:= (_nTemp * (100-_nTemp5))/100 // diferencao do percentual da tabela de vendas
						_aTemp[_nX][09] := "X"
						// Atualizar a ZT7
						For _nZ:=1 to Len(oDlgCar:aCols)
							If !oDlgCar:aCols[_nZ][Len(oDlgCar:aHeader)+1]
								If oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})] == _aTemp[_nX][01]
									oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := _nTemp1
									Exit
								EndIF
							EndIf
						Next
					EndIf
				Next
			Else
				// Rateio por clientes             
				For _nX:=1 to Len(_aTemp)
					_lTemProximo := .F.
					For _nY:=1 to Len(_aTemp)
						If _nY <> _nX .and. _aTemp[_nY][03]==_aTemp[_nX][03] .and. _aTemp[_nY][04]==_aTemp[_nX][04] .and. _aTemp[_nY][05]==_aTemp[_nX][05] .and. _aTemp[_nY][06]==_aTemp[_nX][06] .and. _aTemp[_nY][02]<>_aTemp[_nX][02]
							_lTemProximo 	:= .T.   
							Exit
						EndIF
					Next          
					If !_lTemProximo
						_nTemp := _nTotVlr / Len(_aClientes)
						// Atualizar a ZT7
						For _nZ:=1 to Len(oDlgCar:aCols)
							If !oDlgCar:aCols[_nZ][Len(oDlgCar:aHeader)+1]
								If oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})] == _aTemp[_nX][01]
									oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := _nTemp
									Exit
								EndIF
							EndIf
						Next
					Else
						_nTemp := _nTotVlr / Len(_aClientes)  
						_nReto := ( _nTemp * _aTemp[_nX][10] ) / 100
						_nPeca := _nTemp - _nReto
						
						If _aTemp[_nX][02] == "P"
							// Atualizar a peca na ZA7
							For _nZ:=1 to Len(oDlgCar:aCols)
								If !oDlgCar:aCols[_nZ][Len(oDlgCar:aHeader)+1]
									If oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})] == _aTemp[_nX][01]
										oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := _nReto
										Exit
									EndIF
								EndIf
							Next
							// Atualizar o retorno na ZA7
							For _nY:=1 to len(_aTemp)
								If _nY <> _nX .and. _aTemp[_nY][03]==_aTemp[_nX][03] .and. _aTemp[_nY][04]==_aTemp[_nX][04] .and. _aTemp[_nY][05]==_aTemp[_nX][05] .and. _aTemp[_nY][06]==_aTemp[_nX][06] .and. _aTemp[_nY][02]<>_aTemp[_nX][02] .and. _aTemp[_nY][02]=="R"
									For _nZ:=1 to Len(oDlgCar:aCols)
										If !oDlgCar:aCols[_nZ][Len(oDlgCar:aHeader)+1]
											If oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})] == _aTemp[_nY][01]
												oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := _nPeca
												Exit
											EndIF
										EndIf
									Next
							    EndIF
							Next
						EndIF
						// Encontrar o retorno relacionado na ZA7					
						If _aTemp[_nX][02] == "R"
							// Atualizar a peca na ZA7
							For _nZ:=1 to Len(oDlgCar:aCols)
								If !oDlgCar:aCols[_nZ][Len(oDlgCar:aHeader)+1]
									If oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})] == _aTemp[_nX][01]
										oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := _nPeca
										Exit
									EndIF
								EndIf
							Next
							// Atualizar o retorno na ZA7
							For _nY:=1 to len(_aTemp)
								If _nY <> _nX .and. _aTemp[_nY][03]==_aTemp[_nX][03] .and. _aTemp[_nY][04]==_aTemp[_nX][04] .and. _aTemp[_nY][05]==_aTemp[_nX][05] .and. _aTemp[_nY][06]==_aTemp[_nX][06] .and. _aTemp[_nY][02]<>_aTemp[_nX][02] .and. _aTemp[_nY][02]=="P"
									For _nZ:=1 to Len(oDlgCar:aCols)
										If !oDlgCar:aCols[_nZ][Len(oDlgCar:aHeader)+1]
											If oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})] == _aTemp[_nY][01]
												oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := _nReto
												Exit
											EndIF
										EndIf
									Next
							    EndIF
							Next
						EndIF
					EndIF
				Next
			EndIf
		EndIf
		
		// Frank 02/03/2016 
		_aRet := U_KMMOD2("1", _nX, "2")
		If _aRet[1] == "N" .and. _aRet[3] == "S" // Tratamento especial para rateio line haul

			_nVRFIXO 	:= 0
			_nVRRETO    := 0
			_nTotPeso   := 0
			
			dbSelectArea("ZTC")
			DbOrderNickName("ITUPZTC001")
			dbSeek(xFilial("ZTC")+ZT0->ZT0_CODTAB+ZT0->ZT0_VERTAB+ZT0->ZT0_ITEMTB)
			_aFaixa := {}      
			While !Eof() .and. ZTC_FILIAL+ZTC_TABVEN+ZTC_VERVEN+ZTC_ITTABV == xFilial("ZTC")+ZT0->ZT0_CODTAB+ZT0->ZT0_VERTAB+ZT0->ZT0_ITEMTB
				If ZTC_FAIXAD == 0 .and. ZTC_FAIXAA == 0 .and. ZTC_VRFIXO > 0
					_nVRFIXO := ZTC_VRFIXO

					If ZTC_VRFIXR == 0
						_nVRRETO := (((( ZTC_VRFIXO * ZTC_PERR ) / 100 ) + ZTC_VRFIXO ) - ZTC_VRFIXO)
					Else                                                                             
						_nVRRETO := ZTC_VRFIXR - ZTC_VRFIXO
					EndIF
					
					exit
				EndIF
			    dbSkip()                                           
			EndDo                                                  

			// Calcular as pecas 
			_nTotPeso := 0
			If _nVRFIXO > 0
				For _nX:=1 to Len(oDlgCar:aCols)    
					If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
						If oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})] == "P"
							_nTotPeso += oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_PESO"})]
						EndIF
					EndIf
				Next
				For _nX:=1 to Len(oDlgCar:aCols)    
					If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
						If oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})] == "P"
							oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := (oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_PESO"})] / _nTotPeso) * _nVRFIXO
						EndIF
					EndIf
				Next
			EndIf
				
			// Calcular os retornos
			_nTotPeso := 0
			If _nVRRETO > 0
				For _nX:=1 to Len(oDlgCar:aCols)    
					If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
						If oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})] == "R"
							_nTotPeso += oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_PESO"})]
						EndIF
					EndIf
				Next
				For _nX:=1 to Len(oDlgCar:aCols)    
					If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
						If oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})] == "R"
							oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := (oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_PESO"})] / _nTotPeso) * _nVRRETO
						EndIF
					EndIf
				Next
			EndIF
		
		EndIF
		

		// Adicionar Valor de Picking		-- Cristiam Rossi em 22/02/2016
		For _nX:=1 to Len(oDlgCar:aCols)
			If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
//				If oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})]  == "P"
				If oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_PICKIN"})]  == "S"
					oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})]  += ZT0->ZT0_VRPICK
				endif
			endif
		next
	ElseIf ZT0->ZT0_TIPREG == "5"		

		_dData 		:= oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_DTFIM"})] 		
		_nDia  		:= dow(_dData)
		cTabX  		:= oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TABVEN"})] 
		cVerX  		:= oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_VERVEN"})] 
		cITTX  		:= oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_ITTABV"})] 
		_nVALKM 	:= 0
		_nEKMVEN 	:= 0
		_nVALKMA 	:= 0
		
		// Posicionar na tabela de turnos
		dbSelectArea("ZTE")
		DbOrderNickName("ITUPZTE001")
		ZTE->( dbSeek(xFilial("ZTE")+cTabX+cVerX+cITTX) )
		While !ZTE->(Eof()) .and. ZTE->ZTE_FILIAL+ZTE->ZTE_TABVEN+ZTE->ZTE_VERVEN+ZTE->ZTE_ITTABV == xFilial("ZTE")+cTabX+cVerX+cITTX
			If val(ZTE->ZTE_DIASEM)+1 == _nDia 
				_nVALKM  := ZTE->ZTE_VALKM
				_nVALKMA := ZTE->ZTE_VALKMA
				Exit
			EndIf
			dbSkip()
		EndDo       

		If _nVALKM+_nVALKMA > 0		
		
			For _nX:=1 to Len(oDlgCar:aCols)	
				_nTemp 		:= 0                     
				_nSKMVen    := oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SKMVEN"})]   
				_nEKMVEN    := oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_EKMVEN"})]
		
				If _nEKMVEN == 0  
					_nTemp := _nSKMVen *  _nVALKM			
				Else                                                      
					_nTemp1 := _nSKMVen - _nEKMVEN
					_nTemp2 := _nTemp1 * _nVALKM
					_nTemp3 := _nEKMVEN * _nVALKMA
					_nTemp  := _nTemp3 + _nTemp2
				EndIF
				
				if ZT0->ZT0_DESLOC == "S"		// Carrega Valor de Deslocamento - Cristiam Rossi em 05/02/2016
					_nTemp4 := ZT0->ZT0_VRDES *  gdFieldGet("ZA7_KMDESL" , _nX, .F., oDlgCar:aHeader, oDlgCar:aCols)
					_nTemp += _nTemp4
			    	gdFieldPut("ZA7_VALDES" , _nTemp4, _nX, oDlgCar:aHeader, oDlgCar:aCols)
				else
			    	gdFieldPut("ZA7_VALDES" , 0      , _nX, oDlgCar:aHeader, oDlgCar:aCols)
				endif
				//Pedrassi Begin
				If ZT0->ZT0_TIPLKM == "D" .AND. ZTE->ZTE_VALKM >= 0
					_nTemp	+=	ZT0->ZT0_VALOR 
				EndIf
				//Pedrassi End
				oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := _nTemp
			Next
		Else                                    
			For _nX:=1 to Len(oDlgCar:aCols)	
				oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})] := 0		
			Next
		EndIF		

	endif

	oDlgCar:Refresh()
    		    	
	_VTotal := 0
	For nX := 1 To Len(oDlgCar:aCols)
		If !oDlgCar:aCols[nX][len(oDlgCar:aHeader)+1]    		
			_VTotal += oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})]
		EndIF
	Next nX    


	nValFechT := _VTotal   
	nValFrete := _VTotal   

	oValFechT:Refresh()
	oValFrete:Refresh()

	oDlgCar:oBrowse:Refresh()	
	
Return(_Ret)                  





//********************************************************************************// 	
// Rotina Calculo da Tebela de Compras - padrao GEFCO                             //
// DATA : 22-07-2015                                                              //
//********************************************************************************//
User Function CalcZA7C(_lGerAcols)
Local aColsAux := {}
Local cOrig,cDest,cTpTab,cTpFlu,dDtIni,cViagem
Local cDevEmb,cAdic,cEmerg
Local cTpTran
Local cCodCli	:= 	AllTrim(ZA0->ZA0_CLI)
Local cLojCli	:=  AllTrim(ZA0->ZA0_LOJA)
Local cTabVen	//:=  AllTrim(ZA0->ZA0_TABVEN)
Local cVerTab
Local cTabCom
Local cVerTabC
Local nVALOR,nVEMERG,nPEMERG := 0
Local mSQL       := ""
Local _VTotal 	 := 0
Local _VTotalV 	 := 0 
Local _vFrete 	 := 0
Local _Ret 		 := .t.
Local _cMsgAlert := ""

Local cProjet
Local cObra
Local cSeqC
Local aCamposSim := {}
Local aHeader := {}	
Local aCols := {}

Default _lGerAcols := .F.        

If _lGerAcols // Tratamento especial para compatibilizar a rotina LOCT005

	cProjet := oGetZA7:aCols[oGetZA7:nAT][Ascan(oGetZA7:aHeader,{|x|AllTrim(x[2])=="ZA7_PROJET"})]
	cObra   := oGetZA7:aCols[oGetZA7:nAT][Ascan(oGetZA7:aHeader,{|x|AllTrim(x[2])=="ZA7_OBRA"})]
	cSeqC   := oGetZA7:aCols[oGetZA7:nAT][Ascan(oGetZA7:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCAR"})]

	nStyle:=0
	cAlias   :="ZA6"
	cChave   :=xFILIAL(cAlias)+cProjet+cObra
	cCondicao:='ZA6_FILIAL+ZA6_PROJET+ZA6_OBRA=="'+cChave+'"' 
	nIndice  :=1  
	cFiltro   :='ZA6_FILIAL+ZA6_PROJET+ZA6_OBRA=="'+cChave+'"

	AAdd(aCamposSim,{"ZA6_OBRA"  ,"V"})
	AAdd(aCamposSim,{"ZA6_SEQTRA","V"}) 
	AAdd(aCamposSim,{"ZA6_AS"    ,"V"})
	AAdd(aCamposSim,{"ZA6_CVA"   ,"V"})
	AAdd(aCamposSim,{"ZA6_VIAGEM","V"})
	AAdd(aCamposSim,{"ZA6_TPTRAN","" })
	AAdd(aCamposSim,{"ZA6_EMERGE","" })
	AAdd(aCamposSim,{"ZA6_EMERG2","" })
	AAdd(aCamposSim,{"ZA6_FERIAV","" })     
	AAdd(aCamposSim,{"ZA6_FERIAC","" }) 
	AAdd(aCamposSim,{"ZA6_TIPFLU","" })
	AAdd(aCamposSim,{"ZA6_TRANSP","" })
	AAdd(aCamposSim,{"ZA6_DTINI" ,"" })
	AAdd(aCamposSim,{"ZA6_DTFIM" ,"" })
	AAdd(aCamposSim,{"ZA6_ORIGEM","" })
	AAdd(aCamposSim,{"ZA6_MUNORI","" })
	AAdd(aCamposSim,{"ZA6_ESTORI","" })
	AAdd(aCamposSim,{"ZA6_DESTIN","" })
	AAdd(aCamposSim,{"ZA6_MUNDES","" })
	AAdd(aCamposSim,{"ZA6_ESTDES","" })
	AAdd(aCamposSim,{"ZA6_DESCRO","" })
	AAdd(aCamposSim,{"ZA6_MONITO","" })
	AAdd(aCamposSim,{"ZA6_OBSVIA","" })
	AAdd(aCamposSim,{"ZA6_CCGEFC","" })
	AAdd(aCamposSim,{"ZA6_CCCLIE","" })
	AAdd(aCamposSim,{"ZA6_OI"	 ,"" })
	AAdd(aCamposSim,{"ZA6_CONTA" ,"" })
	AAdd(aCamposSim,{"ZA6_TIPDES","" })
	AAdd(aCamposSim,{"ZA6_REFGEF","" })
	AAdd(aCamposSim,{"ZA6_OBSTRA","" })
	AAdd(aCamposSim,{"ZA6_CLIORI","" })
	AAdd(aCamposSim,{"ZA6_LOJORI","" })
	AAdd(aCamposSim,{"ZA6_NOMORI","" })  
	AAdd(aCamposSim,{"ZA6_CEPORI","" })
	AAdd(aCamposSim,{"ZA6_CLIDES","" })
	AAdd(aCamposSim,{"ZA6_LOJDES","" })
	AAdd(aCamposSim,{"ZA6_NOMDES","" })
	AAdd(aCamposSim,{"ZA6_CGCDES","" })
	AAdd(aCamposSim,{"ZA6_ENDDES","" })
	AAdd(aCamposSim,{"ZA6_INSDES","" })
	AAdd(aCamposSim,{"ZA6_BAIDES","" })
	AAdd(aCamposSim,{"ZA6_MUNDE2","" })
	AAdd(aCamposSim,{"ZA6_ESTDE2","" })
	AAdd(aCamposSim,{"ZA6_CEPDES","" })
	AAdd(aCamposSim,{"ZA6_CONDES","" })
	AAdd(aCamposSim,{"ZA6_DEPDES","" })
	AAdd(aCamposSim,{"ZA6_EMADES","" })
	AAdd(aCamposSim,{"ZA6_DDDDES","" })
	AAdd(aCamposSim,{"ZA6_TELDES","" })
	AAdd(aCamposSim,{"ZA6_PEDCLI","" })
	AAdd(aCamposSim,{"ZA6_SOLICT","" })
	AAdd(aCamposSim,{"ZA6_CONPAG","" })
	AAdd(aCamposSim,{"ZA6_TPFRET","" })
	AAdd(aCamposSim,{"ZA6_TIPPAG","" })
	AAdd(aCamposSim,{"ZA6_TIPLKM","" })
	AAdd(aCamposSim,{"ZA6_SKMVEN","" })
	AAdd(aCamposSim,{"ZA6_EKMVEN","" })
	AAdd(aCamposSim,{"ZA6_SKMCOM","" })
	AAdd(aCamposSim,{"ZA6_EKMCOM","" })
	AAdd(aCamposSim,{"ZA6_THORAS","" })
	AAdd(aCamposSim,{"ZA6_TABVEN","V"}) 
	AAdd(aCamposSim,{"ZA6_VERVEN","V"}) 
	AAdd(aCamposSim,{"ZA6_ITTABV","V"})     
	AAdd(aCamposSim,{"ZA6_TABCOM","V"})                                
	AAdd(aCamposSim,{"ZA6_VERCOM","V"})
	AAdd(aCamposSim,{"ZA6_ITTABC","V"})      

	aHeader:=fHeader(aCamposSim)
	aCols:=fCols(aHeader,cAlias,nIndice,cChave,cCondicao,cFiltro)
                                                                                                                                            
	oDlgTra := oCalcX():new()

	oDlgTra:aCols 	:= aClone(aCols)
	oDlgTra:aHeader := aClone(aHeader)
	oDlgTra:nAt		:= 1
	
	
	//oDlgTra:=MsNewGetDados():New(2000,2000,2500,2500,nStyle,"AllwaysFalse()","AllwaysFalse()","+ZA6_SEQTRA",,,9999,,,.t.,oDlg,aHeader,aCols)

	aCamposSim := {}
	nStyle:=0

	cAlias   :="ZA7"
	cChave   :=xFILIAL(cAlias)+cProjet+cObra
	cCondicao:='ZA7_FILIAL+ZA7_PROJET+ZA7_OBRA=="'+cChave+'"'
	nIndice  :=1  
	cFiltro  :='ZA7_FILIAL+ZA7_PROJET+ZA7_OBRA=="'+cChave+'"'


	AAdd(aCamposSim,{"ZA7_OBRA"  ,"V"})
	AAdd(aCamposSim,{"ZA7_SEQTRA","V"})
	AAdd(aCamposSim,{"ZA7_SEQCAR","V"})
	AAdd(aCamposSim,{"ZA7_EHTERC","" })
	AAdd(aCamposSim,{"ZA7_QUANT" ,"" })
	AAdd(aCamposSim,{"ZA7_EMERGE","" })
	AAdd(aCamposSim,{"ZA7_EMERG2","" })
	AAdd(aCamposSim,{"ZA7_FILREG","" })
	AAdd(aCamposSim,{"ZA7_DEVEMB","" })
	AAdd(aCamposSim,{"ZA7_SEQCOL","" })
	AAdd(aCamposSim,{"ZA7_ADICIO","" })
	AAdd(aCamposSim,{"ZA7_CODCLI","" })
	AAdd(aCamposSim,{"ZA7_LOJCLI","" })
	AAdd(aCamposSim,{"ZA7_DESCLI","" })
	AAdd(aCamposSim,{"ZA7_MUNICI","" })
	AAdd(aCamposSim,{"ZA7_DESMUN","" })
	AAdd(aCamposSim,{"ZA7_UFCARG","" })
	AAdd(aCamposSim,{"ZA7_CLIDES","" })
	AAdd(aCamposSim,{"ZA7_LOJDES","" })
	AAdd(aCamposSim,{"ZA7_NOMDES","" })
	AAdd(aCamposSim,{"ZA7_CARGA" ,"" })
	AAdd(aCamposSim,{"ZA7_PICKIN","" })
	AAdd(aCamposSim,{"ZA7_QTD" 	 ,"" })
	AAdd(aCamposSim,{"ZA7_JUNTO" ,"" })
	AAdd(aCamposSim,{"ZA7_COMP"  ,"" })
	AAdd(aCamposSim,{"ZA7_LARG"  ,"" })
	AAdd(aCamposSim,{"ZA7_ALTU"  ,"" })
	AAdd(aCamposSim,{"ZA7_PESO"  ,"" })
	AAdd(aCamposSim,{"ZA7_VRCARG","" })
	AAdd(aCamposSim,{"ZA7_CARENC","" })
	AAdd(aCamposSim,{"ZA7_TIPCAR","" })
	AAdd(aCamposSim,{"ZA7_CAREND","" })
	AAdd(aCamposSim,{"ZA7_TPCARD","" })
	AAdd(aCamposSim,{"ZA7_FORMAS","" })
	AAdd(aCamposSim,{"ZA7_VRMOB" ,"" })
	AAdd(aCamposSim,{"ZA7_PERADV","" })
	AAdd(aCamposSim,{"ZA7_VALADV","" })
	AAdd(aCamposSim,{"ZA7_INCADV","" })
	AAdd(aCamposSim,{"ZA7_RESPC" ,"" })
	AAdd(aCamposSim,{"ZA7_DTCAR" ,"" })
	AAdd(aCamposSim,{"ZA7_HRCAR" ,"" })
	AAdd(aCamposSim,{"ZA7_RESPD" ,"" })
	AAdd(aCamposSim,{"ZA7_DTDES" ,"" })
	AAdd(aCamposSim,{"ZA7_HRDES" ,"" })
	AAdd(aCamposSim,{"ZA7_CCGEFC","" })
	AAdd(aCamposSim,{"ZA7_CCCLIE","" })
	AAdd(aCamposSim,{"ZA7_OI"	 ,"" })
	AAdd(aCamposSim,{"ZA7_CONTA" ,"" })
	AAdd(aCamposSim,{"ZA7_TIPDES","" })
	AAdd(aCamposSim,{"ZA7_REFGEF","" })
	AAdd(aCamposSim,{"ZA7_OBS" 	 ,"" })
	AAdd(aCamposSim,{"ZA7_REVNAS","" })
	AAdd(aCamposSim,{"ZA7_AS" 	 ,"" })
	AAdd(aCamposSim,{"ZA7_VIAGEM","" })
	AAdd(aCamposSim,{"ZA7_VALESC","" })
	AAdd(aCamposSim,{"ZA7_VALBAL","" })
	AAdd(aCamposSim,{"ZA7_VALAJU","" })
	AAdd(aCamposSim,{"ZA7_VALCD" ,"" })
	AAdd(aCamposSim,{"ZA7_KMDESL","" })
	AAdd(aCamposSim,{"ZA7_VALDES","" })
	AAdd(aCamposSim,{"ZA7_VRDESC","" })
	AAdd(aCamposSim,{"ZA7_VALLIC","" })
	AAdd(aCamposSim,{"ZA7_VALOR" ,"" })
	AAdd(aCamposSim,{"ZA7_CUSTO" ,"" })
	AAdd(aCamposSim,{"ZA7_SKMVEN","" })
	AAdd(aCamposSim,{"ZA7_EKMVEN","" })
	AAdd(aCamposSim,{"ZA7_SKMCOM","" })
	AAdd(aCamposSim,{"ZA7_EKMCOM","" }) 

	aHeader:=fHeader(aCamposSim)
	aCols:=fCols(aHeader,cAlias,nIndice,cChave,cCondicao,cFiltro)

	oDlgCar := oCalcX():new()
	
	oDlgCar:aCols 	:= aClone(aCols)
	oDlgCar:aHeader := aClone(aHeader)
	oDlgCar:nAt		:= 1
	
	
EndIF


	cViagem   := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_OBRA"})]  //aColsAux[Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_OBRA"})] 
	cTpTrans  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TPTRAN"})] //aColsAux[Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TPTRAN"})] 
	cTpFluxo  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TIPFLU"})] //aColsAux[Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TIPFLU"})]
	dDtIni	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_DTINI" })] //aColsAux[Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_DTINI"})] 
	cEmerg	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_EMERGE"})] //aColsAux[Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_EMERGE"})] 
	cTabCom	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TABCOM"})]
	cVerTabC  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_VERCOM"})] //aColsAux[Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_VERVEN"})] 								
	cITTabC	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_ITTABC"})] //aColsAux[Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_VERVEN"})] 								
	_Orig	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_ORIGEM"})]
	_Dest	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_DESTIN"})]					
//	cTES	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TES"})]
	cCONPAG	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_CONPAG"})]
	cTPFRET	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TPFRET"})]
	cCCGEFC	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_CCGEFC"})]
	cCCCLIE	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_CCCLIE"})]
	cOI	  	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_OI" })]
	cCONTA 	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_CONTA"})]
	cTIPDES   := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TIPDES"})]                   
	cTpTran	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|Alltrim(x[2])=="ZA6_TRANSP"})] 

/*
	if Len(oDlgCon:aCols) ==0 
		Alert("Na aba conjunto transportador não esta preenchida.")
		Return
	Else
		cTpTran := oDlgCon:aCols[1][Ascan(oDlgCon:aHeader,{|x|Alltrim(x[2])=="ZAE_TRANSP"})]
		if Empty(Alltrim(cTpTran))
			Alert("Na aba conjunto transportador não esta preenchida.")
			Return
		Endif	 
	Endif
*/

	if Empty(Alltrim(cTpTran))
		Alert("O Campo Transportador não esta preenchida.")
		Return(.f.)
	Endif

/*
	if Empty(Alltrim(cTES))
		Alert("O Campo Tipo Saida não esta preenchida.")
		Return(.f.)
	Endif
*/            

	if Empty(Alltrim(cCONPAG))
		Alert("O campo Condições de Pagamento não esta preenchida.")
		Return(.f.)
	Endif
	            
	if Empty(Alltrim(cTPFRET))
		Alert("O campo Tipo de Frete não esta preenchida.")
		Return(.f.)
	Endif

	if Empty(Alltrim(cCCGEFC))
		Alert("O campo Centro de Custo GEFCO  não esta preenchida.")
		Return(.f.)
	Endif
            
	if Empty(Alltrim(cCCCLIE)) .and. SuperGetMv("IT_VALCLI",,"90136500") == cCodCli+cLojCli
		Alert("O campo Centro de Custo Cliente  não esta preenchida.")
		Return(.f.)
	Endif

	if Empty(Alltrim(cOI)) .and. SuperGetMv("IT_VALCLI",,"90136500") == cCodCli+cLojCli
		Alert("O campo Ordem Interna não esta preenchida.")
		Return(.f.)
	Endif

	if Empty(Alltrim(cCONTA)) .and. SuperGetMv("IT_VALCLI",,"90136500") == cCodCli+cLojCli
		Alert("O campo Conta PSA não esta preenchida.")
		Return(.f.)
	Endif

	if Empty(Alltrim(cTIPDES))
		Alert("O campo Tipo Despesa GEFCO não esta preenchida.")
		Return(.f.)
	Endif
            
/*
	DbSelectArea("ZAE")
	DbSetOrder(1)
	DbSeek(xFilial("ZAE")+ZA0->ZA0_PROJET+cVIAGEM)
	cTpTran := ZAE->ZAE_TRANSP
*/
             

	cTabX	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TABCOM"})]
	cVerX	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_VERCOM"})]
	cITTX	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_ITTABC"})]
	dbSelectArea("ZT1")
	dbSetOrder(1)                                                                                
//	dbSeek( xFilial("ZT1") + ZT0->ZT0_TABCOM + ZT0->ZT0_VERCOM)
	dbSeek( xFilial("ZT1") + cTabX + cVerX)
	_lAchou := .F.
	While !Eof() .and. ZT1_FILIAL+ZT1_CODTAB+ZT1_VERTAB == xFilial("ZT1")+cTabX + cVerX
		If ZT1_TIPTAB == cTpTrans .and. ZT1_ITEMTB == cITTX .and. ZT1_CODCLI == cCodCli .and. ZT1_LOJCLI == cLojCli
			_lAchou := .T.
			Exit
		EndIF
		dbSkip()
	EndDo
	If !_lAchou
		MsgStop("Tabela de Compras não identificada.","Atenção!")
		Return .F.
	EndIF
    
	aArea := GetArea()

	If ZT1->ZT1_TIPREG < "4" .and. _lGerAcols // modelo de calculo utilizado na coleta LOCT005
	
		For nPos := 1 to Len(oGetZA7:aCols)
			if ( _SqCarPos := Ascan(oGetZA7:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCAR"}) ) == 0
				_SqCarPos := Ascan(oGetZA7:aHeader,{|x|AllTrim(x[2])=="ZL8_SEQCAR"})
			endif
	
			cSeqC := oGetZA7:aCols[nPos][_SqCarPos]

			IF SELECT("TZA7") > 0
				dbSelectArea("TZA7")
				dbCloseArea()
			Endif
	                    
			mSQL := "SELECT ZA6_TPTRAN,ZA6_TIPFLU,ZA6_DTINI,ZA6_EMERG2,ZA6_TABCOM,ZA6_VERCOM, "
			mSQL += " ZA6_ITTABC,ZA6_ORIGEM,ZA6_DESTIN,ZA6_TIPDES,ZA6_TRANSP,ZA7_DEVEMB, ZA7_EMERG2, "
			mSQL += " ZA7_ADICIO,ZA7_MUNICI,ZA7.R_E_C_N_O_ AS ZA7RECNO,ZA6.R_E_C_N_O_ AS ZA6RECNO " 
			mSQL += " FROM "+RetSQLName("ZA7")+" ZA7 ,"+RetSQLName("ZA6")+" ZA6 "
			mSQL += " WHERE ZA7.D_E_L_E_T_ =' ' AND ZA6.D_E_L_E_T_ =' ' "
			mSQL += " AND ZA6_FILIAL='"+xFilial("ZA6")+"' AND ZA7_FILIAL='"+xFilial("ZA7")+"'"
			mSQL += " AND ZA6_PROJET=ZA7_PROJET AND ZA6_AS=ZA7_AS AND ZA6_AS<>''"
			mSQL += " AND ZA7_PROJET='"+cProjet+"' AND ZA7_OBRA='"+cObra+"'"
			mSQL += " AND ZA7_SEQCAR='"+cSeqC+"'"
	
			dbUseArea( .T., "TOPCONN", TCGENQRY(,,mSQL), "TZA7", .F., .T. )
					
			if ! EOF()
				cDevEmb   := TZA7->ZA7_DEVEMB
				cAdic     := TZA7->ZA7_ADICIO
				cMunic    := TZA7->ZA7_MUNICI
	     		cTpTrans  := TZA7->ZA6_TPTRAN
				cTpFluxo  := TZA7->ZA6_TIPFLU
				dDtIni	  := TZA7->ZA6_DTINI
				cEmerg	  := TZA7->ZA7_EMERG2
				cTabCom	  := TZA7->ZA6_TABCOM
				cVerTabC  := TZA7->ZA6_VERCOM
				cITTabC	  := TZA7->ZA6_ITTABC
				cOrig	  := TZA7->ZA6_ORIGEM
				cDest	  := TZA7->ZA6_DESTIN
				cTIPDES   := TZA7->ZA6_TIPDES                  
				
				cFrota    := oDlgFro:aCols[oDlgFro:nAt][Ascan(oDlgFro:aHeader,{|x|AllTrim(x[2])=="ZL9_FROTA"})]
				cMotori	  := oDlgFro:aCols[oDlgFro:nAt][Ascan(oDlgFro:aHeader,{|x|AllTrim(x[2])=="ZL9_MOTORI"})] 
				
				nRecZA7	  := TZA7->ZA7RECNO
				nRecZA6   := TZA7->ZA6RECNO
			endif
	
			IF SELECT("TZA7") > 0
				dbSelectArea("TZA7")
				dbCloseArea()
				RestArea( aArea )
			Endif
			
			IF Empty(AllTrim(cFrota))
				Alert("Favor Campo Frota precisa ser preenchido")
			    Return(.f.)
			Endif 
			IF Empty(AllTrim(cMotori))
				Alert("Favor Campo Motorista precisa ser preenchido")
			    Return(.f.)
			Endif 
	
			cForn	  := Posicione("DA4",1,xFilial("DA4") + cMotori , "DA4_FORNEC")	
			cTpTran	  := Posicione("DA3",1,xFilial("DA3") + cFrota  , "DA3_TIPVEI")
	
			mSQL := "SELECT ZA2_CODIGO,ZA2_DESCRI,ZA2_ESTADO,ZA2_CDPAIS,ZA2_CODMUN, "
			mSQL += " ZT1_CODTAB,ZT1_VERTAB,ZT1_ITEMTB,ZT1_CUSTO,ZT1_TIPVEI,ZT1_PEMERG,ZT1_TIPTAB "
			mSQL += " FROM "+RetSQLName("ZA2")+" ZA2 LEFT JOIN "+RetSQLName("ZT1")+" ZT1 "
			mSQL += " ON ZT1_CODORI=ZA2_CODIGO "
			mSQL += " AND ZT1_CODCLI='"+cCodCLi+"' AND ZT1_LOJCLI='"+cLojCli+"' "
			mSQL += " AND ZT1_TIPTAB='"+cTpTrans+"' AND ZT1_TIPFLU='"+cTpFluxo+"' "
			mSQL += " AND ZT1_CODORI='"+cOrig+"' AND ZT1_CODDES='"+cDest+"' "
			mSQL += " WHERE ZT1_FILIAL='"+xFilial("ZT1")+"' AND ZT1.D_E_L_E_T_=' ' "
			mSQL += " AND ZA2_FILIAL='"+xFilial("ZA2")+"' AND ZA2.D_E_L_E_T_=' ' "
			If valtype(dDtIni) <> "D"
				mSQL += " AND ZT1_INIVIG <='"+ dDtIni +"' AND ( ZT1_FIMVIG >='"+ dDtIni +"'  OR ZT1_FIMVIG ='' ) "
			Else
				mSQL += " AND ZT1_INIVIG <='"+ dtos(dDtIni) +"' AND ( ZT1_FIMVIG >='"+ dtos(dDtIni) +"'  OR ZT1_FIMVIG ='' ) "
			EndIF
			mSQL += " AND ZT1_MSBLQL='2' " //AND ZA2_CODIGO='"+_Munic+"' "
			mSQL += " AND ZT1_TIPVEI='"+cTpTran+"' AND ZT1_CODFOR='"+cForn+"'" 
			mSQL += " GROUP BY ZA2_CODIGO,ZA2_DESCRI,ZA2_ESTADO,ZA2_CDPAIS,ZA2_CODMUN, ZT1_CODTAB,ZT1_VERTAB, "
			mSQL += " ZT1_ITEMTB,ZT1_CUSTO,ZT1_TIPVEI, ZT1_PEMERG, ZT1_TIPTAB "
	
			dbUseArea( .T., "TOPCONN", TCGENQRY(,,mSQL), "TZA2", .F., .T. )
	
			if ! EOF()
				_vFrete := TZA2->ZT1_CUSTO
				nPEMERG := TZA2->ZT1_PEMERG 
				cTabCOM := TZA2->ZT1_CODTAB
				cVerTabC:= TZA2->ZT1_VERTAB 
				cITTabC := TZA2->ZT1_ITEMTB
				cTpTrans:= TZA2->ZT1_TIPTAB
				cRegra  := "1"
			endif
	
			IF SELECT("TZA2") > 0
				dbSelectArea("TZA2")
				dbCloseArea()
				RestArea( aArea )
			Endif               
	
			if Empty(AllTrim(cTabCOM))
				Return .t.
			Endif
			
			IF cEmerg =="N" .AND. cDevEmb == "P" .AND. cAdic =="N" // ITEM 1.1.1
				nValor := _vFrete
	    	    ZA7->(dbGoto(nRecZA7)) 
				RecLock("ZA7",.F.)
				ZA7->ZA7_CUSTO := nValor
				ZA7->(MsUnLock())
		
	    	    ZA6->(dbGoto(nRecZA6)) 
				RecLock("ZA6",.F.)
				ZA6->ZA6_TABCOM := cTabCom
				ZA6->ZA6_VERCOM := cVerTabC
			  	ZA6->ZA6_ITTABC := cITTabC 
				ZA6->(MsUnLock())
		
			Elseif cEmerg =="S" .AND. cDevEmb == "P" .AND. cAdic =="N"  // ITEM 1.1.2
				nValor :=_vFrete+(_vFrete * (nPEMERG/100))
	    	    ZA7->(dbGoto(nRecZA7)) 
				RecLock("ZA7",.F.)
				ZA7->ZA7_CUSTO := nValor
				ZA7->(MsUnLock())
	
	    	    ZA6->(dbGoto(nRecZA6)) 
				RecLock("ZA6",.F.)
				ZA6->ZA6_TABCOM := cTabCom
				ZA6->ZA6_VERCOM := cVerTabC
			  	ZA6->ZA6_ITTABC := cITTabC 
				ZA6->(MsUnLock())
	
			Elseif cEmerg =="N"  .AND. cDevEmb == "P" .AND. cAdic =="S"	 // ITEM 1.1.3				
				
				IF SELECT("TZTB") > 0
					dbSelectArea("TZTB")
					dbCloseArea()
				Endif
	  
				mSQL := "SELECT ZTB_TABCOM,ZTB_VERTAB,ZTB_CODMUN,ZTB_CUSTO "    
				mSQL += "FROM "+RetSQLName("ZTB")+" ZTB "
				mSQL += " WHERE ZTB_FILIAL ='"+xFilial("ZTB")+"' AND D_E_L_E_T_=' ' "
				mSQL += " AND ZTB_TABCOM='"+cTabCom+"' "
				mSQL += " AND ZTB_VERTAB='"+cVerTabC+"' AND ZTB_CODMUN='"+cMunic+"'"
				mSQL += " AND ZTB_ITTABC='"+cITTabC+"' AND ZTB_MSBLQL='2' "  
				
				dbUseArea( .T., "TOPCONN", TCGENQRY(,,mSQL), "TZTB", .F., .T. )
	
				IF TZTB->(!EoF())
					nValor := TZTB->ZTB_CUSTO
	    	    
		    	    ZA7->(dbGoto(nRecZA7)) 
					RecLock("ZA7",.F.)
					ZA7->ZA7_CUSTO := nValor
					ZA7->(MsUnLock())
						
		    	    ZA6->(dbGoto(nRecZA6)) 
					RecLock("ZA6",.F.)
					ZA6->ZA6_TABCOM := cTabCom
					ZA6->ZA6_VERCOM := cVerTabC
				  	ZA6->ZA6_ITTABC := cITTabC 
					ZA6->(MsUnLock())
				Endif
	
	            TZTB->(dbCloseArea())
							
			Elseif cEmerg =="S"  .AND. cDevEmb == "P" .AND. cAdic =="S"	 // ITEM 1.1.4
										
				IF SELECT("TZTB") > 0
					dbSelectArea("TZTB")
					TZTB->(dbCloseArea())
				Endif
	
				mSQL := "SELECT ZTB_TABCOM,ZTB_VERTAB,ZTB_CODMUN,ZTB_CUSTO "    
				mSQL += " FROM "+RetSQLName("ZTB")+" ZTB "
				mSQL += " WHERE ZTB_FILIAL ='"+xFilial("ZTB")+"' AND D_E_L_E_T_=' ' "
				mSQL += " AND ZTB_TABCOM='"+cTabCom+"' "
				mSQL += " AND ZTB_VERTAB='"+cVerTabC+"' AND ZTB_CODMUN='"+cMunic+"'"
				mSQL += " AND ZTB_ITTABC='"+cITTabC+"' AND ZTB_MSBLQL='2' " 						 
							
				dbUseArea( .T., "TOPCONN", TCGENQRY(,,mSQL), "TZTB", .F., .T. )
			
				IF TZTB->(!EoF())
					nValor := TZTB->ZTB_CUSTO +(TZTB->ZTB_CUSTO*nPEMERG/100)
	
		    	    ZA7->(dbGoto(nRecZA7)) 
					RecLock("ZA7",.F.)
						ZA7->ZA7_CUSTO := nValor
					ZA7->(MsUnLock())
	
		    	    ZA6->(dbGoto(nRecZA6)) 
					RecLock("ZA6",.F.)
						ZA6->ZA6_TABCOM := cTabCom
						ZA6->ZA6_VERCOM := cVerTabC
					  	ZA6->ZA6_ITTABC := cITTabC 
						ZA6->(MsUnLock())
	
				Endif
	
				TZTB->(dbCloseArea())
	
			Elseif cEmerg =="N" .AND. cDevEmb == "R" .AND. cAdic =="N" // ITEM 1.1.6
				nValor := _vFrete
	    	    ZA7->(dbGoto(nRecZA7)) 
				RecLock("ZA7",.F.)
				ZA7->ZA7_CUSTO := 0//nValor
				ZA7->(MsUnLock())
	
	    	    ZA6->(dbGoto(nRecZA6)) 
				RecLock("ZA6",.F.)
				ZA6->ZA6_TABCOM := cTabCom
				ZA6->ZA6_VERCOM := cVerTabC
			  	ZA6->ZA6_ITTABC := cITTabC 
				ZA6->(MsUnLock())
		
			Elseif cEmerg =="S"  .AND. cDevEmb == "R" .AND. cAdic =="N"   // ITEM 1.1.7
				if nVEMERG == 0 
					nValor := _vFrete+(_vFrete * (nPEMERG/100))
				Else
					nValor := nVEMERG
				Endif
	
	    	    ZA7->(dbGoto(nRecZA7)) 
				RecLock("ZA7",.F.)
				ZA7->ZA7_CUSTO := 0//nValor
				ZA7->(MsUnLock())
	
	    	    ZA6->(dbGoto(nRecZA6)) 
				RecLock("ZA6",.F.)
				ZA6->ZA6_TABCOM := cTabCom
				ZA6->ZA6_VERCOM := cVerTabC
			  	ZA6->ZA6_ITTABC := cITTabC 
				ZA6->(MsUnLock())
		
			Elseif cEmerg =="N"  .AND. cDevEmb == "R" .AND. cAdic =="S"	 // ITEM 1.1.8				
				nValor := _vFrete+(_vFrete * (nPEMERG/100))
			    ZA7->(dbGoto(nRecZA7)) 
				RecLock("ZA7",.F.)
				ZA7->ZA7_CUSTO := 0//nValor
				ZA7->(MsUnLock())
		
			    ZA6->(dbGoto(nRecZA6)) 
				RecLock("ZA6",.F.)
				ZA6->ZA6_TABCOM := cTabCom
				ZA6->ZA6_VERCOM := cVerTabC
				ZA6->ZA6_ITTABC := cITTabC 
				ZA6->(MsUnLock())
		
			Elseif cEmerg =="S"  .AND. cDevEmb == "R" .AND. cAdic =="S"	 // ITEM 1.1.9
		
				If SELECT("TZTB") > 0
					dbSelectArea("TZTB")
					TZTB->(dbCloseArea())
				Endif
		
				mSQL := "SELECT ZTB_TABCOM,ZTB_VERTAB,ZTB_CODMUN,ZTB_CUSTO "    
				mSQL += " FROM "+RetSQLName("ZTB")+" ZTB "
				mSQL += " WHERE ZTB_FILIAL ='"+xFilial("ZTB")+"' AND D_E_L_E_T_=' ' "
				mSQL += " AND ZTB_TABCOM='"+cTabCom+"' "
				mSQL += " AND ZTB_VERTAB='"+cVerTabC+"' AND ZTB_CODMUN='"+cMunic+"'"
				mSQL += " AND ZTB_ITTABC='"+cITTabC+"' AND ZTB_MSBLQL='2' " 						 
					
				dbUseArea( .T., "TOPCONN", TCGENQRY(,,mSQL), "TZTB", .F., .T. )
				
				IF TZTB->(!EoF())
					nValor := TZTB->ZTB_CUSTO +(TZTB->ZTB_CUSTO*nPEMERG/100)
				    ZA7->(dbGoto(nRecZA7)) 
					RecLock("ZA7",.F.)
					ZA7->ZA7_CUSTO := nValor
					ZA7->(MsUnLock())
		
	    	    	ZA6->(dbGoto(nRecZA6)) 
					RecLock("ZA6",.F.)
					ZA6->ZA6_TABCOM := cTabCom
					ZA6->ZA6_VERCOM := cVerTabC
					ZA6->ZA6_ITTABC := cITTabC 
					ZA6->(MsUnLock())
			
				Endif
		
				TZTB->(dbCloseArea())
			Endif
		Next nPos
	
	EndIF             
	
	
	If ZT1->ZT1_TIPREG < "4" .and. !_lGerAcols	// Modelo utilizado no agendamento
	
		_cMsgAlert := ""

		For nX := 1 To Len(oDlgCar:aCols)
	
			cDevEmb := oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})]
			cAdic   := oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_ADICIO"})]
			_Munic  := oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_MUNICI"})]
	
			if Select("TZA2") > 0
				TZA2->(dbCloseArea())
			endif
	
			mSQL := " SELECT ZA2_CODIGO,ZA2_DESCRI,ZA2_ESTADO,ZA2_CDPAIS,ZA2_CODMUN, "
			mSQL += " ZT1_CODTAB,ZT1_VERTAB,ZT1_ITEMTB,ZT1_CUSTO,ZT1_TIPVEI,ZT1_PEMERG "
			mSQL += " FROM "+RetSQLName("ZA2")+" ZA2 INNER JOIN "+RetSQLName("ZT1")+" ZT1 "
			mSQL += " ON ZT1_CODORI=ZA2_CODIGO AND ZT1_CODTAB='"+cTabCOM+"' AND ZT1_VERTAB='"+cVerTabC+"'" 
			If	!Empty(cITTabC)
				mSQL += " AND ZT1_ITEMTB='"+cITTabC+"' "
			EndIf
			mSQL += " AND ZT1_CODCLI='"+cCodCLi+"' AND ZT1_LOJCLI='"+cLojCli+"' "
			mSQL += " AND ZT1_TIPTAB='"+cTpTrans+"' AND ZT1_TIPFLU='"+cTpFluxo+"' "
			//mSQL += " AND ZT1_CODORI='"+_Orig+"' AND ZT1_CODDES='"+_Dest+"' "
			mSQL += " WHERE ZT1_FILIAL='"+xFilial("ZT1")+"' AND ZT1.D_E_L_E_T_=' ' "
			mSQL += " AND ZA2_FILIAL='"+xFilial("ZA2")+"' AND ZA2.D_E_L_E_T_=' ' "
			If Valtype(dDtIni) == "D"
				mSQL += " AND ZT1_INIVIG <='"+DtoS(dDtIni)+"' AND ( ZT1_FIMVIG >='"+DtoS(dDtIni)+"'  OR ZT1_FIMVIG ='' ) "
			Else
				mSQL += " AND ZT1_INIVIG <='"+dDtIni+"' AND ( ZT1_FIMVIG >='"+dDtIni+"'  OR ZT1_FIMVIG ='' ) "
			EndIf
			If	Empty(cITTabC)
				mSQL += " AND ZT1_TIPVEI='"+cTpTran+"'"	
			EndIf
			mSQL += " AND ZT1_MSBLQL='2' " //AND ZA2_CODIGO='"+_Munic+"' "
			mSQL += " GROUP BY ZA2_CODIGO,ZA2_DESCRI,ZA2_ESTADO,ZA2_CDPAIS,ZA2_CODMUN, ZT1_CODTAB,ZT1_VERTAB, "
			mSQL += " ZT1_ITEMTB,ZT1_CUSTO,ZT1_TIPVEI, ZT1_PEMERG "
			dbUseArea( .T., "TOPCONN", TCGENQRY(,,mSQL), "TZA2", .F., .T. )
	
			dbSelectArea("TZA2")
			if ! EOF()
				_vFrete := TZA2->ZT1_CUSTO
				nPEMERG := TZA2->ZT1_PEMERG 
				cITTabC := TZA2->ZT1_ITEMTB
				oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_ITTABC"})] := cITTabC
			else
	//			Aviso("A T E N Ç Ã O","O valor informado não consta na Tabela de Custo de Frete "+Trim(Str(nX))+" da aba Coletas" ,{"Ok"}, 2)
				//_cMsgAlert += "O valor informado não consta na Tabela de Compra do Frete "+Trim(Str(nX))+" da aba Coletas" + CRLF
				Loop
			endif
			TZA2->(dbCloseArea())
	        
			cEmerg := oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_EMERG2"})]
	
			if cDevEmb == "R"
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]	:= 0
			elseIf cEmerg =="N" .AND. cDevEmb == "P" .AND. cAdic =="N" // ITEM 1.1.1
				nValor := _vFrete
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]	:= nValor
			Elseif cEmerg =="S" .AND. cDevEmb == "P" .AND. cAdic =="N"  // ITEM 1.1.2
				nValor :=_vFrete+(_vFrete * (nPEMERG/100))
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]	:= nValor
			Elseif cEmerg =="N"  .AND. cDevEmb == "P" .AND. cAdic =="S"	 // ITEM 1.1.3				
				
				IF SELECT("TZTB") > 0
					dbSelectArea("TZTB")
					TZTB->(dbCloseArea())
				Endif
	  
				mSQL := "SELECT ZTB_TABCOM,ZTB_VERTAB,ZTB_CODMUN,ZTB_CUSTO "    
				mSQL += "FROM "+RetSQLName("ZTB")+" ZTB "
				mSQL += " WHERE ZTB_FILIAL ='"+xFilial("ZTB")+"' AND D_E_L_E_T_=' ' "
				mSQL += " AND ZTB_TABCOM='"+cTabCom+"' "
				mSQL += " AND ZTB_VERTAB='"+cVerTabC+"' AND ZTB_CODMUN='"+_Munic+"'"
				mSQL += " AND ZTB_ITTABC='"+cITTabC+"' AND ZTB_MSBLQL != '1' "  
				
				dbUseArea( .T., "TOPCONN", TCGENQRY(,,mSQL), "TZTB", .F., .T. )
	
				dbSelectArea("TZTB")
				IF ! EoF()
					nValor := TZTB->ZTB_CUSTO
					oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]	:= nValor
				else
	//				Aviso("A T E N Ç Ã O","O valor adicional informado não consta na Tabela de Compra do Frete "+Trim(Str(nX))+" da aba Coletas" ,{"Ok"}, 2)
					//_cMsgAlert += "O valor adicional informado não consta na Tabela de Compra do Frete "+Trim(Str(nX))+" da aba Coletas" + CRLF
					oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]	:= 0
					Loop
				Endif
				TZTB->(dbCloseArea())
	
			Elseif cEmerg =="S"  .AND. cDevEmb == "P" .AND. cAdic =="S"	 // ITEM 1.1.4
								
				IF SELECT("TZTB") > 0
					dbSelectArea("TZTB")
					TZTB->(dbCloseArea())
				Endif
	
				mSQL := "SELECT ZTB_TABCOM,ZTB_VERTAB,ZTB_CODMUN,ZTB_CUSTO "    
				mSQL += " FROM "+RetSQLName("ZTB")+" ZTB "
				mSQL += " WHERE ZTB_FILIAL ='"+xFilial("ZTB")+"' AND D_E_L_E_T_=' ' "
				mSQL += " AND ZTB_TABCOM='"+cTabCom+"' "
				mSQL += " AND ZTB_VERTAB='"+cVerTabC+"' AND ZTB_CODMUN='"+_Munic+"'"
				mSQL += " AND ZTB_ITTABC='"+cITTabC+"' AND ZTB_MSBLQL='2' " 						 
	
				dbUseArea( .T., "TOPCONN", TCGENQRY(,,mSQL), "TZTB", .F., .T. )
	
				dbSelectArea("TZTB")
				IF TZTB->(!EoF())
					nValor := TZTB->ZTB_CUSTO +(TZTB->ZTB_CUSTO*nPEMERG/100)
					oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]	:= nValor
				Endif
	
				TZTB->(dbCloseArea())
			Elseif cEmerg =="N" .AND. cDevEmb == "R" .AND. cAdic =="N" // ITEM 1.1.6
				nValor := _vFrete
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]	:= nValor
			Elseif cEmerg =="S"  .AND. cDevEmb == "R" .AND. cAdic =="N"   // ITEM 1.1.7
				if nVEMERG == 0 
					nValor :=_vFrete+(_vFrete * (nPEMERG/100))
				Else
					nValor := nVEMERG
				Endif
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]	:= nValor
	
			Elseif cEmerg =="N"  .AND. cDevEmb == "R" .AND. cAdic =="S"	 // ITEM 1.1.8				
				nValor := _vFrete+(_vFrete * (nPEMERG/100))
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]	:= nValor
			Elseif cEmerg =="S"  .AND. cDevEmb == "R" .AND. cAdic =="S"	 // ITEM 1.1.9
	
				IF SELECT("TZTB") > 0
					dbSelectArea("TZTB")
					TZTB->(dbCloseArea())
				Endif
	
				mSQL := "SELECT ZTB_TABCOM,ZTB_VERTAB,ZTB_CODMUN,ZTB_CUSTO "    
				mSQL += " FROM "+RetSQLName("ZTB")+" ZTB "
				mSQL += " WHERE ZTB_FILIAL ='"+xFilial("ZTB")+"' AND D_E_L_E_T_=' ' "
				mSQL += " AND ZTB_TABCOM='"+cTabCom+"' "
				mSQL += " AND ZTB_VERTAB='"+cVerTabC+"' AND ZTB_CODMUN='"+_Munic+"'"
				mSQL += " AND ZTB_ITTABC='"+cITTabC+"' AND ZTB_MSBLQL='2' " 						 
				dbUseArea( .T., "TOPCONN", TCGENQRY(,,mSQL), "TZTB", .F., .T. )
	
				dbSelectArea("TZTB")
				IF TZTB->(!EoF())
					nValor := TZTB->ZTB_CUSTO +(TZTB->ZTB_CUSTO*nPEMERG/100)
					oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]	:= nValor
				Else
	//				Aviso("A T E N Ç Ã O","O valor adicional informado não consta na Tabela de Compra do Frete "+Trim(Str(nX))+" da aba Coletas" ,{"Ok"}, 2)
	//				_cMsgAlert += "O valor adicional informado não consta na Tabela de Compra do Frete "+Trim(Str(nX))+" da aba Coletas" + CRLF
					oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]	:= 0
					Loop
				Endif
	
				TZTB->(dbCloseArea())
			Endif
	
		Next nX
	
	
	EndIF	
	


	If ZT1->ZT1_TIPREG == "4"

		// Passo 1 - Calcular o ZA7_CUSTO para todas as Pecas
		// --------------------------------------------------
		cTabX	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TABCOM"})] 
		cVerX	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_VERCOM"})] 
		cITTX	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_ITTABC"})] 
		dbSelectArea("ZT1")
		dbSetOrder(1)                                                                                
		dbSeek( xFilial("ZT1") + cTabX + cVerX)
		_lAchou := .F.
		While !Eof() .and. ZT1_FILIAL+ZT1_CODTAB+ZT1_VERTAB == xFilial("ZT1") + cTabX + cVerX
			If ZT1_TIPTAB == cTpTrans .and. ZT1_ITEMTB == cITTX .and. ZT1_CODCLI == cCodCli .and. ZT1_LOJCLI == cLojCli
				_lAchou := .T.
				Exit
			EndIF
			dbSkip()
		EndDo


		For _nX:=1 to Len(oDlgCar:aCols)	
			If !oDlgCar:aCols[_nX][len(oDlgCar:aHeader)+1]    		
				If oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})]=="P"
                    
					// PASSO 1 - Encontrar o ZA7_CUSTO para a PECA posicionada
					// -------------------------------------------------------
					_nKm := U_KMMOD2("2", _nX, "1") //oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SKMCOM"})]
					
					dbSelectArea("ZTD")
					DbOrderNickName("ITUPZTD001")
					_nVlrK := 0
					dbSeek(xFilial("ZTD")+ZT1->ZT1_CODTAB+ZT1->ZT1_VERTAB+ZT1->ZT1_ITEMTB)
					_aFaixa := {}      
					_lVRFIXO := .F.
					While !Eof() .and. ZTD_FILIAL+ZTD_TABCOM+ZTD_VERCOM+ZTD_ITTABC == xFilial("ZTD")+ZT1->ZT1_CODTAB+ZT1->ZT1_VERTAB+ZT1->ZT1_ITEMTB
						aadd(_aFaixa, {ZTD_FAIXAD, ZTD_FAIXAA, ZTD_VRKMON, ZTD_VRFIXO, ZTD_VREONE })
						If ZTD_VRFIXO > 0                                      
							_lVRFIXO := .T.
						EndIF
					    dbSkip()
					EndDo                                                                 
					              
					_nVlrK := 0
					_lSai := .F.
					For _nZ:=1 to len(_aFaixa)
						If _aFaixa[_nZ][01]==0 .and. _aFaixa[_nZ][02]==0
							_nVlrK := _aFaixa[_nZ][03] * _nKm // VRKMON
							Exit  
						Else                         
						
							_lSai := .F.
							If !_lVRFIXO                            
								If _aFaixa[_nZ][01] <= _nKm .and. _aFaixa[_nZ][2] >= _nKm
									_nVlrK := _aFaixa[_nZ][03] * _nKm // VRKMON
									Exit
								EndIF
							Else  
								If _aFaixa[_nZ][01] <= _nKm .and. _aFaixa[_nZ][2] >= _nKm
									If _nZ == len(_aFaixa)
										_nTemp  := _nKm - _aFaixa[_nZ-1][2]
										_nTemp1 := _nTemp * _aFaixa[_nZ][05]
										_nTemp2 := _nKm - _nTemp
										For _nT:=1 to len(_aFaixa)
											If _aFaixa[_nT][01] <= _nTemp2 .and. _aFaixa[_nT][2] >= _nTemp2
												_nVlrK := _nTemp1 + _aFaixa[_nT][04]
												_lSai := .T.
												Exit
											EndIF
										Next
										If _lSai
											Exit
										EndIF
									Else
										_nVlrK := _aFaixa[_nZ][04] // VRFIXO e dentro de uma faixa que não é a ultima
										_lSai := .T.
										Exit
									EndIf
								EndIF						
							EndIf
							If _lSai
								Exit
							EndIF

						EndIf
							
					Next            
					
					oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})] := _nVlrK
					 
			    EndIF
			EndIF
		Next				 
		
		// PASSO 2 - Calcular o ZA7_CUSTO para os retornos.
		// ------------------------------------------------
		For _nX:=1 to Len(oDlgCar:aCols)	
			If !oDlgCar:aCols[_nX][len(oDlgCar:aHeader)+1] 
				If oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})]=="R"		
					_nVlrK := 0   		
					
					// Encontrar a peca correspondente ao registro de retorno.
					// -------------------------------------------------------
					_cCliTmp := oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CODCLI"})]
					_cLojTmp := oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJCLI"})]
					_cCliDes := oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CLIDES"})]
					_cLojDes := oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJDES"})]
					_nKmPeca := 0
					_nVlPeca := 0
					For _nY:=1 to Len(oDlgCar:aCols)	
						If !oDlgCar:aCols[_nY][len(oDlgCar:aHeader)+1] 
							If oDlgCar:aCols[_nY][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})]=="P"		
								If oDlgCar:aCols[_nY][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CODCLI"})]==_cCliTmp
									If oDlgCar:aCols[_nY][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJCLI"})]==_cLojTmp
										If oDlgCar:aCols[_nY][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CLIDES"})]==_cCliDes
											If oDlgCar:aCols[_nY][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJDES"})]==_cLojDes
												_nKmPeca := oDlgCar:aCols[_nY][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SKMCOM"})]
												_nVlPeca := oDlgCar:aCols[_nY][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]
												Exit
											EndIf
										EndIF
									EndIF
								EndIF
							EndIF
						EndIF
					Next       
					
					// Encontrar as faixas validas.
					// ----------------------------
					dbSelectArea("ZTD")
					DbOrderNickName("ITUPZTD001")
					_nVlrK := 0
					dbSeek(xFilial("ZTD")+ZT1->ZT1_CODTAB+ZT1->ZT1_VERTAB+ZT1->ZT1_ITEMTB)
					_aFaixa := {}      
					_lVRFIXO := .F.
					While !Eof() .and. ZTD_FILIAL+ZTD_TABCOM+ZTD_VERCOM+ZTD_ITTABC == xFilial("ZTD")+ZT1->ZT1_CODTAB+ZT1->ZT1_VERTAB+ZT1->ZT1_ITEMTB
						aadd(_aFaixa, {ZTD_FAIXAD, ZTD_FAIXAA, ZTD_VRKMON, ZTD_VRFIXO, ZTD_VREONE, ZTD_VRKMRO, ZTD_PEKMRO, ZTD_VRFIXR, ZTD_PERR, ZTD_VREROU })
					    dbSkip()
					EndDo

					_nTemp := 0					
					For _nZ:=1 to len(_aFaixa)
						_lSai := .F.
						// Calculo por faixa = 0
						// -----------------------------------------
						If _aFaixa[_nZ][01]==0 .and. _aFaixa[_nZ][02]==0
							If _aFaixa[_nZ][06] > 0 // VRKMRO
								_nTemp := (_nKmPeca * _aFaixa[_nZ][06]) - _nVlPeca
							Else
								_nTemp := (_nKmPeca * ((  (_aFaixa[_nZ][03]*_aFaixa[_nZ][07])/100   )+_aFaixa[_nZ][03])) - _nVlPeca
							EndIF

							_nKm := U_KMMOD2("2", _nX, "1") //oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SKMCOM"})]
							If _aFaixa[_nZ][06] > 0 // VRKMRO
								_nTemp := (_nKm * _aFaixa[_nZ][06]) + _nTemp
							Else
								_nTemp := (_nKm * (( (_aFaixa[_nZ][03]*_aFaixa[_nZ][07])/100  )+_aFaixa[_nZ][03])) + _nTemp
							EndIF

							Exit  
						EndIF  
						// Calculo por faixa preenchida
						// ----------------------------
						// CASO 1 - o valor do KM se enquadra em uma faixa sem ser a ultima
						// ----------------------------------------------------------------
						_nKm := U_KMMOD2("2", _nX, "1") //oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SKMCOM"})]
						If _aFaixa[_nZ][01] <= _nKm .and. _aFaixa[_nZ][2] >= _nKm .and. _nZ < len(_aFaixa)
							If _aFaixa[_nZ][04] > 0 // Valor fixo
								If _aFaixa[_nZ][08] > 0 // VRFIXR
									_nTemp := _aFaixa[_nZ][08] - _aFaixa[_nZ][04]
									Exit
								Else
									_nTemp := (_aFaixa[_nZ][04] * _aFaixa[_nZ][09])/100
									Exit
								EndIF
							Else     
								If _aFaixa[_nZ][06] == 0 // VRKMRO
									_nTemp1 := ((_aFaixa[_nZ][03] * _aFaixa[_nZ][07])/100)+_aFaixa[_nZ][03]
									_nTemp2 := (_nKmPeca * _nTemp1) - _nVlPeca
									_nTemp  := (_nTemp1 * _nKm) + _nTemp2
									Exit
								Else                                     
									_nTemp1 := _aFaixa[_nZ][06] * _nKmPeca
									_nTemp2 := _nTemp1 - _nVlPeca
									_nTemp  := (_nKm * _aFaixa[_nZ][06]) + _nTemp2
									Exit
								EndIF 
							EndIF
						EndIF
						// CASO 2 - o valor do KM se enquadra na ultima faixa 
						// ----------------------------------------------------------------
						If _aFaixa[_nZ][01] <= _nKm .and. _aFaixa[_nZ][2] >= _nKm .and. _nZ == len(_aFaixa)
							If _aFaixa[_nZ-1][04] > 0 // vlr fixo
								// posicionar na faixa anterior
								If _aFaixa[_nZ-1][08] > 0 // VRFIXR
									_nTemp1 := _aFaixa[_nZ-1][08] - _aFaixa[_nZ-1][04]
									_nTemp2 := _nKmPeca - _aFaixa[len(_aFaixa)-1][02]
									_nTemp3 := _nTemp2 * _aFaixa[_nZ][05] // VREONE
									_nTemp4 := _nTemp2 * _aFaixa[_nZ][10] // VREROU
									_nTemp5 := _nTemp4 - _nTemp3
									_nTemp6 := _nKm - _aFaixa[len(_aFaixa)-1][02]
									_nTemp  := _nTemp6 * _aFaixa[_nZ][10] 
									_nTemp  := _nTemp + _nTemp5 + _nTemp1
									exit
								Else    
									_nTemp1 := (_aFaixa[_nZ-1][04]*_aFaixa[_nZ-1][09])/100
									_nTemp2 := _nKmPeca - _aFaixa[len(_aFaixa)-1][02]
									_nTemp3 := _nTemp2 * _aFaixa[_nZ][05] // VREONE
									_nTemp4 := _nTemp2 * _aFaixa[_nZ][10] // VREROU
									_nTemp5 := _nTemp4 - _nTemp3
									_nTemp6 := _nKm - _aFaixa[len(_aFaixa)-1][02]
									_nTemp  := _nTemp6 * _aFaixa[_nZ][10] 
									_nTemp  := _nTemp + _nTemp5 + _nTemp1
									exit
								EndIF
								
							Else                               
								
								_nTemp1 := _nKmPeca * _aFaixa[_nZ][06]  // _VRKMRO
								_nTemp2 := _nTemp1 - _nVlPeca
								_nTemp3 := _nKm * _aFaixa[_nZ][06]
								_nTemp  := _nTemp3 + _nTemp2
								Exit
								
							EndIF
						EndIF
						
					Next
					
					oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})] := _nTemp
					
				EndIf 
			EndIF
		Next


		//------------ Considerar Cálculo Emergencial
		// Posicionar na tabela de compras
		cTabX	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TABCOM"})] 
		cVerX	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_VERCOM"})] 
		cITTX	  := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_ITTABC"})] 

		nPEMERG := 0
		nPERFER := 0

		ZT1->( dbSetOrder(1) )
		ZT1->( dbSeek( xFilial("ZT1") + cTabX + cVerX) )
		While ! ZT1->( Eof() ) .and. ZT1->(ZT1_FILIAL+ZT1_CODTAB+ZT1_VERTAB) == xFilial("ZT1") + cTabX + cVerX
			If ZT1->ZT1_TIPTAB == cTpTrans .and. ZT1->ZT1_ITEMTB == cITTX .and. ZT1->ZT1_CODCLI == cCodCli .and. ZT1->ZT1_LOJCLI == cLojCli
				nPEMERG := ZT1->ZT1_PEMERG
				nPERFER := ZT1->ZT1_PERFER
				Exit
			EndIF
			ZT1->( dbSkip() )
		EndDo
		
		if gdFieldGet("ZA6_FERIAC", oDlgTra:nAt, .F., oDlgTra:aHeader, oDlgTra:aCols ) != "S"	// não vai tratar o % Feriado
			nPERFER := 0
		endif

		For _nX := 1 to Len(oDlgCar:aCols)
			if ! gdDeleted( _nX, oDlgCar:aHeader, oDlgCar:aCols)
				if gdFieldGet("ZA7_EMERGE", _nX, .F., oDlgCar:aHeader, oDlgCar:aCols ) == "S" .and. nPEMERG > 0	// se for SIM, adicionar % de Emergência
					nValor := gdFieldGet("ZA7_CUSTO", _nX, .F., oDlgCar:aHeader, oDlgCar:aCols ) 
					nValor += nValor * nPEMERG / 100
			    	gdFieldPut("ZA7_CUSTO", nValor, _nX, oDlgCar:aHeader, oDlgCar:aCols)
				endif

				if nPERFER > 0	// se for SIM, adicionar % de Feriado
					nValor := gdFieldGet("ZA7_CUSTO", _nX, .F., oDlgCar:aHeader, oDlgCar:aCols ) 
					nValor += nValor * nPERFER / 100
			    	gdFieldPut("ZA7_CUSTO", nValor, _nX, oDlgCar:aHeader, oDlgCar:aCols)
				endif
			endif
		next                                                                         

		_aRet := U_KMMOD2("1", _nX, "2")
		If _aRet[2] == "S" // Tratamento especial para rateio por KM/PESO
		
			// Tratamento especial para rateio por KM ou PESO
			// ----------------------------------------------------------------------------------------------------------------

			// Detalhamento do Retorno da função KMMOD2
			// cVenda , cCompra , cRatVen , cRatCom , nPerV , nPerC
			// cVenda  == "S" representa que havera rateio para o calculo da venda
			// cCompra == "S" representa que havera rateio para o calculo da compra
			// cRatVen == "N" representa que o rateio da venda sera por KM
			// cRatVen == "S" representa que o rateio da venda sera por PESO
			// cRatCom == "N" representa que o rateio da compra sera por KM
			// cRatCom == "S" representa que o rateio da compra sera por PESO
			// nPerV          representa o % do rateio da embalagem para venda
			// nPerC	      representa o % do rateio da embalagem para compra                  
	
			// PASSO 1 - Validacoes especiais para realizar o calculo
			// ----------------------------------------------------------------------------------------------------------------
			_lErro := .F.
			For _nX:=1 to Len(oDlgCar:aCols)    
				If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
					_aRet := U_KMMOD2("2", _nX, "2")
					If _aRet[1] == "S" // todos os rateios devem ter o seqcol preenchido
						If empty(oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})])
							_lErro := .T.
						EndIF
					EndIF
				EndIF
			Next
			If _lErro
				MsgStop("As sequencias da coleta nao foram preenchidas.","Atenção!")
				Return _Ret
			EndIf
			// ----------------------------------------------------------------------------------------------------------------
			// PASSO 2 - Encontrar o total de KM de todas as coletas associadas que tenha o campo SEQCOL preenchido
			_nTotKM := 0
	
			dbSelectArea("ZTD")
			DbOrderNickName("ITUPZTD001")
			_nVlrK := 0
			dbSeek(xFilial("ZTD")+ZT1->ZT1_CODTAB+ZT1->ZT1_VERTAB+ZT1->ZT1_ITEMTB)
			_aFaixa := {}      
			While !Eof() .and. ZTD_FILIAL+ZTD_TABCOM+ZTD_VERCOM+ZTD_ITTABC == xFilial("ZTD")+ZT1->ZT1_CODTAB+ZT1->ZT1_VERTAB+ZT1->ZT1_ITEMTB
				aadd(_aFaixa, {ZTD_FAIXAD, ZTD_FAIXAA, ZTD_VRKMRO, ZTD_VRFIXO })
			    dbSkip()
			EndDo                                                                 
	
			For _nX:=1 to Len(oDlgCar:aCols)    
				If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
					_aRet := U_KMMOD2("2", _nX, "2")
					If _aRet[1] == "S" 
						If !empty(oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})])
							_nTotKm += oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SKMCOM"})]					
						EndIF
					EndIf
				EndIF
			Next                     
	
			// Calculo do valor total
			_nTotVlr  := 0
			If _nTotKm > 0
				_nTotVlr  := 0
				For _nZ:=1 to len(_aFaixa)
					If _aFaixa[_nZ][01]==0 .and. _aFaixa[_nZ][02]==0
						_nTotVlr  := _aFaixa[_nZ][03] * _nTotKm // VRKMON
						Exit  
					EndIf                         
					If _aFaixa[_nZ][01] <= _nTotKm .and. _aFaixa[_nZ][2] >= _nTotKm
						If _aFaixa[_nZ][04] > 0
							_nTotVlr  := _aFaixa[_nZ][04] // VRFIXO e dentro de uma faixa que não é a ultima
						Else
							_nTotVlr  := _aFaixa[_nZ][03] * _nTotKm
						EndIF
						Exit
					EndIF						
				Next
			EndIF                     
			
			// Encontrar quantos clientes diferentes existem na ZA7 e destes quais possuem pecas.
			// Quando ocorre uma coleta que existe apenas o retorno (não há uma peça), a rotina precisa identificar todas os clientes
			// diferentes da ZA7 que possuem pecas para realizar o rateio. Uma peça por cliente.
			_aPecas := {}
			For _nX:=1 to Len(oDlgCar:aCols)    
				If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
					_aRet := U_KMMOD2("2", _nX, "2")
					If _aRet[1] == "S"              
						If !empty(oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})]) .and. oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})] == "P"
							If Len(_aPecas) == 0
								aadd(_aPecas,{oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CODCLI"})]+;
								              oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJCLI"})],;
								              oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})],;
								              0})
							Else 
	
	        					_cTemp  := oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CODCLI"})] + oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJCLI"})]
	        					_lAchou := If( aScan(_aPecas, {|x| x[1] == _cTemp }) > 0, .T., .F.)
	
								If !_lAchou
									aadd(_aPecas,{oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CODCLI"})]+;
								              oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJCLI"})],;
								              oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})],;
								              0})						
							    EndIF
							EndIF
						EndIF					
					EndIf
				EndIF
			Next
			
			// PASSO 2.1 - Encontrar o total de PESO de todas as coletas associadas que tenha o campo SEQCOL preenchido
			_nTotP := 0
			For _nX:=1 to Len(oDlgCar:aCols)    
				If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
					_aRet := U_KMMOD2("2", _nX, "2")
					If _aRet[1] == "S"              
						// Totalizar o peso somente para as pecas
						If !empty(oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})]) //.and. oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})] == "P"
							_nTotP += oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_PESO"})]
						EndIF
					EndIf
				EndIF
			Next
			// --------------------------------------------------------------------------------------------------------------
			// -------------------------------------------------------------------------------------------------------------
			// PASSO 3 - Geracao do temporario para realizacao dos calculos
			_aTemp := {}
			For _nX:=1 to Len(oDlgCar:aCols)    
				If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
					_aRet := U_KMMOD2("2", _nX, "2")
					If _aRet[1] == "S"
						If !empty(oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})])
							aadd(_aTemp, {oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})],;
							 			  oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})],;
							 			  oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CODCLI"})],;
							 			  oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJCLI"})],;   
							 			  oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CLIDES"})],;       
							 			  oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_LOJDES"})],;
							 			  oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SKMCOM"})],;
							 			  0,;
							 			  " ",;
							 			  _aRet[6],;
							 			  oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_PESO"})] })
						EndIF
					EndIf
				EndIF
			Next
			_aTemp := aSort(_aTemp,,,{|x,y| x[1] < y[1] })
			// -------------------------------------------------------------------------------------------------------------
			// Calcular primeiro todos as coletas que nao tem rateio pela tabela de vendas
			For _nX:=1 to Len(_aTemp)
				_lTemProximo := .F.
				_nTemp       := 0
				If _aTemp[_nX][09] == "X"
					Loop
				EndIF
				For _nY:=1 to Len(_aTemp)
					If _nY <> _nX .and. _aTemp[_nY][03]==_aTemp[_nX][03] .and. _aTemp[_nY][04]==_aTemp[_nX][04] .and. _aTemp[_nY][05]==_aTemp[_nX][05] .and. _aTemp[_nY][06]==_aTemp[_nX][06] .and. _aTemp[_nY][02]<>_aTemp[_nX][02]
						_lTemProximo 	:= .T.   
						Exit
					EndIF
				Next
				If !_lTemProximo .and. _aTemp[_nX][2] == "P"           
					// Se nao tem proximo e for peca nao ha rateio, porem se for Retorno há o rateio por pecas por cliente diferente
					_aRet := U_KMMOD2("2", _nX, "2")
					If _aRet[3] == "N" // rateio por km
						_aTemp[_nX][08] := (_aTemp[_nX][07]/_nTotKM) *  _nTotVlr 
					Else // rateio por peso
						_aTemp[_nX][08] := (_aTemp[_nX][11]/_nTotP) *  _nTotVlr 
					EndIF
					// Atualizar a ZA7
					For _nZ:=1 to Len(oDlgCar:aCols)
						If !oDlgCar:aCols[_nZ][Len(oDlgCar:aHeader)+1]
							If oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})] == _aTemp[_nX][01]
								oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})] := _aTemp[_nX][08]
							EndIF
						EndIf
					Next
					_aTemp[_nX][09] := "X" 
				EndIF
				
				If !_lTemProximo .and. _aTemp[_nX][2] == "R"           
					// Se nao tem proximo e for peca nao ha rateio, porem se for Retorno há o rateio por pecas por cliente diferente
					_aRet := U_KMMOD2("2", _nX, "2")
					If _aRet[3] == "N" // rateio por km              
						If len(_aPecas) > 0
							_aTemp[_nX][08] := ( ( (_aTemp[_nX][07] / _nTotKM) * _nTotVlr ) * _aTemp[_nX][10] ) / 100
							_nDiferenca     := ( ( _aTemp[_nX][07] / _nTotKM ) * _nTotVlr) - _aTemp[_nX][08]
							
							For _nH:=1 to Len(_aPecas) 
								_aPecas[_nH][03] += _nDiferenca / len(_aPecas)
							Next
							
						Else                                                                                
							_aTemp[_nX][08] := ( _aTemp[_nX][07] / _nTotKM ) *  _nTotVlr
						EndIF
					Else // rateio por peso
						If len(_aPecas) > 0
							_aTemp[_nX][08] := ( ( ( _aTemp[_nX][11] / _nTotP ) *  _nTotVlr ) * _aTemp[_nX][10] ) / 100
							_nDiferenca     := ( ( _aTemp[_nX][11] / _nTotP ) * _nTotVlr ) - _aTemp[_nX][08]
							
							For _nH:=1 to Len(_aPecas) 
								_aPecas[_nH][03] += _nDiferenca/len(_aPecas)
							Next
							
						Else                                                                                
							_aTemp[_nX][08] := ( _aTemp[_nX][11] / _nTotP ) *  _nTotVlr
						EndIF
						
					EndIF                             
					// Atualizar a ZA7
					For _nZ:=1 to Len(oDlgCar:aCols)
						If !oDlgCar:aCols[_nZ][Len(oDlgCar:aHeader)+1]
							If oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})] == _aTemp[_nX][01]
								oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})] := _aTemp[_nX][08]
							EndIF
						EndIf
					Next
					_aTemp[_nX][09] := "X" 
				EndIF
				
			Next
			// ---------------------------------------------------------------------------------------------------------------
			// Calcular os que possuem rateio
			For _nX:=1 to Len(_aTemp)
				If _aTemp[_nX][09] == "X" // ja processado
					Loop
				EndIf
	
				// Encontrar o rateio de todas as coletas associadas
				// -----------------------------------------------------------------------------------------------------------
				_lTemProximo := .F.
				_nTemp       := 0
				_aRet := U_KMMOD2("2", _nX, "2")
				If _aRet[3] == "N" // rateio por km
					_nTemp4      := _aTemp[_nX][07]
				Else // rateio por peso
					_nTemp4      := _aTemp[_nX][11]			
				EndIf			
				// acumular em _nTemp4 o total dos rateios
				_aTemp[_nX][09] := "X"              
				For _nY:=1 to Len(_aTemp)
					If _nY <> _nX .and. _aTemp[_nY][03]==_aTemp[_nX][03] .and. _aTemp[_nY][04]==_aTemp[_nX][04] .and. _aTemp[_nY][05]==_aTemp[_nX][05] .and. _aTemp[_nY][06]==_aTemp[_nX][06] .and. _aTemp[_nY][02]<>_aTemp[_nX][02] .and. _aTemp[_nY][09] == " "
						_aRet := U_KMMOD2("2", _nY, "2")
						If _aRet[3] == "N" // rateio por km
							_nTemp4 		+= _aTemp[_nY][07]
						Else // rateio por peso
							_nTemp4 		+= _aTemp[_nY][11]
						EndIf
						_aTemp[_nY][09] := "X"
						Exit
					EndIF
				Next
				// -----------------------------------------------------------------------------------------------------------
				// Encontrar o rateio por igual para todas as coletas
				_aRet := U_KMMOD2("2", _nY, "2")
				If _aRet[3] == "N" // rateio por km
					_nTemp := (_nTemp4/_nTotKM) *  _nTotVlr 
				Else // rateio por peso
					_nTemp := (_nTemp4/_nTotP) *  _nTotVlr 			
				EndIF                                               
				// -----------------------------------------------------------------------------------------------------------
				// CASO 1 - estamos no retorno
				// Aplicar o percentual da ZT0 no retorno                                   
				_nTemp5 := 0
				If _aTemp[_nX][02] == "R"
					_nTemp1 := (_nTemp * _aTemp[_nX][10])/100 // % da tabela de venda para retorno
					If _aTemp[_nX][10] == 0
						MsgStop("O calculo foi executado com um erro, pois o % de retorno da tabela de vendas esta zerado.","Atenção!")
					EndIf           
					_nTemp5 := _aTemp[_nX][10] // armazenar o % aplicado para realizar a diferenca na peca
					// Atualizar a ZT7                 
					For _nZ:=1 to Len(oDlgCar:aCols)
						If !oDlgCar:aCols[_nZ][Len(oDlgCar:aHeader)+1]
							If oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})] == _aTemp[_nX][01]
								oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})] := _nTemp1
								
								// Verificar se existe um rateio de pecas para aglutinar no valor
								For _nH:=1 to Len(_aPecas)
									If oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})] == _aPecas[_nH][02] .and. oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})] == "P"
										oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})] += _aPecas[_nH][03]
										Exit
									EndIF
								Next
								
							EndIF
						EndIf
					Next
					// Procurar pela peca e aplicar o percentual da diferencao da ZT0
					For _nY:=1 to Len(_aTemp)
						If _nY <> _nX .and. _aTemp[_nY][03]==_aTemp[_nX][03] .and. _aTemp[_nY][04]==_aTemp[_nX][04] .and. _aTemp[_nY][05]==_aTemp[_nX][05] .and. _aTemp[_nY][06]==_aTemp[_nX][06] .and. _aTemp[_nY][02]=="P"
							_nTemp1	:= (_nTemp * (100-_nTemp5))/100 // diferenca do percentual da tabela de vendas
							_aTemp[_nY][09] := "X"
							// Atualizar a ZT7
							For _nZ:=1 to Len(oDlgCar:aCols)
								If !oDlgCar:aCols[_nZ][Len(oDlgCar:aHeader)+1]
									If oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})] == _aTemp[_nY][01]
										oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})] := _nTemp1
										
										// Verificar se existe um rateio de pecas para aglutinar no valor
										For _nH:=1 to Len(_aPecas)
											If oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})] == _aPecas[_nH][02] .and. oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})] == "P"
												oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})] += _aPecas[_nH][03]
												Exit
											EndIF
										Next
										
										Exit
									EndIF
								EndIf
							Next
							Exit
						EndIF
					Next     
				Else // é uma peca      
					// ---------------------------------------------------------------------------------------------------------
					// procurar pelo retorno
					_nTemp5 := 0
					For _nY:=1 to Len(_aTemp)
						If _nY <> _nX .and. _aTemp[_nY][03]==_aTemp[_nX][03] .and. _aTemp[_nY][04]==_aTemp[_nX][04] .and. _aTemp[_nY][05]==_aTemp[_nX][05] .and. _aTemp[_nY][06]==_aTemp[_nX][06] .and. _aTemp[_nY][02]=="R"
							_nTemp1 := (_nTemp * _aTemp[_nY][10])/100 // % da tabela de venda para retorno
							_nTemp5 := _aTemp[_nY][10]
							_aTemp[_nY][10] := "X"
							// Atualizar a ZT7                 
							For _nZ:=1 to Len(oDlgCar:aCols)
								If !oDlgCar:aCols[_nZ][Len(oDlgCar:aHeader)+1]
									If oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})] == _aTemp[_nY][01]
										oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})] := _nTemp1
										
										// Verificar se existe um rateio de pecas para aglutinar no valor
										For _nH:=1 to Len(_aPecas)
											If oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})] == _aPecas[_nH][02] .and. oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})] == "P"
												oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})] += _aPecas[_nH][03]
												Exit
											EndIF
										Next
										
										Exit
									EndIF
								EndIf
							Next
							Exit
						EndIF  
					Next
					// processar a peca           
					_nTemp1	:= (_nTemp * (100-_nTemp5))/100 // diferencao do percentual da tabela de vendas
					_aTemp[_nX][09] := "X"
					// Atualizar a ZT7
					For _nZ:=1 to Len(oDlgCar:aCols)
						If !oDlgCar:aCols[_nZ][Len(oDlgCar:aHeader)+1]
							If oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})] == _aTemp[_nX][01]
								oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})] := _nTemp1
								
								// Verificar se existe um rateio de pecas para aglutinar no valor
								For _nH:=1 to Len(_aPecas)
									If oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SEQCOL"})] == _aPecas[_nH][02] .and. oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})] == "P"
										oDlgCar:aCols[_nZ][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})] += _aPecas[_nH][03]
										Exit
									EndIF
								Next
								
								Exit
							EndIF
						EndIf
					Next
				EndIf
				
			Next
		EndIf
		
		// Frank 02/03/2016 
		_aRet := U_KMMOD2("1", _nX, "2")
		If _aRet[1] == "N" .and. _aRet[3] == "S" // Tratamento especial para rateio line haul

			_nVRFIXO 	:= 0
			_nVRRETO    := 0
			_nTotPeso   := 0


			dbSelectArea("ZTD")
			DbOrderNickName("ITUPZTD001")
			dbSeek(xFilial("ZTD")+ZT1->ZT1_CODTAB+ZT1->ZT1_VERTAB+ZT1->ZT1_ITEMTB)
			While !Eof() .and. ZTD_FILIAL+ZTD_TABCOM+ZTD_VERCOM+ZTD_ITTABC == xFilial("ZTD")+ZT1->ZT1_CODTAB+ZT1->ZT1_VERTAB+ZT1->ZT1_ITEMTB
				
				If ZTD_FAIXAD == 0 .and. ZTD_FAIXAA == 0 .and. ZTD_VRFIXO > 0
					_nVRFIXO := ZTD_VRFIXO

					If ZTD_VRFIXR == 0
						_nVRRETO := (((( ZTD_VRFIXO * ZTD_PERR ) / 100 ) + ZTD_VRFIXO ) - ZTD_VRFIXO)
					Else                                                                             
						_nVRRETO := ZTD_VRFIXR - ZTD_VRFIXO
					EndIF
					
					exit
				EndIF
				
			    dbSkip()
			EndDo

			// Calcular as pecas     
			_nTotPeso := 0
			If _nVRFIXO > 0
				For _nX:=1 to Len(oDlgCar:aCols)    
					If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
						If oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})] == "P"
							_nTotPeso += oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_PESO"})]
						EndIF
					EndIf
				Next
				For _nX:=1 to Len(oDlgCar:aCols)    
					If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
						If oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})] == "P"
							oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})] := (oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_PESO"})] / _nTotPeso) * _nVRFIXO
						EndIF
					EndIf
				Next
			EndIf
				
			// Calcular os retornos  
			_nTotPeso := 0
			If _nVRRETO > 0
				For _nX:=1 to Len(oDlgCar:aCols)    
					If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
						If oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})] == "R"
							_nTotPeso += oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_PESO"})]
						EndIF
					EndIf
				Next
				For _nX:=1 to Len(oDlgCar:aCols)    
					If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
						If oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})] == "R"
							oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})] := (oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_PESO"})] / _nTotPeso) * _nVRRETO
						EndIF
					EndIf
				Next
			EndIF
		
		EndIF
		
		// Tratamento especiao para aglutinar o valor total do custo na primeira linha do ZA7
		// Frank Zwarg Fuga 08-03-2016
		// ---------------------------------------------------------------------------------------------------
		_nTotal := 0
		For _nX:=1 to Len(oDlgCar:aCols)
			If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
				_nTotal += oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]
			EndIf
		Next
		_lPrimeira := .T.		    
		For _nX:=1 to Len(oDlgCar:aCols)
			If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
				If _lPrimeira
					oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})] := _nTotal
					_lPrimeira := .F.
				Else
					oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})] := 0
				EndIF
			EndIf
		Next                                                                                         
		// ---------------------------------------------------------------------------------------------------                                                                                                      
		
		
		
	ElseIf ZT1->ZT1_TIPREG == "5"		

		_dData 		:= oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_DTFIM"})] 		
		_nDia  		:= dow(_dData)
		cTabX  		:= oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TABCOM"})] 
		cVerX  		:= oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_VERCOM"})] 
		cITTX  		:= oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_ITTABC"})] 
		_nVALKM 	:= 0
		_nEKMVEN 	:= 0
		_nVALKMA 	:= 0
		
		// Posicionar na tabela de turnos
		dbSelectArea("ZTF")
		DbOrderNickName("ITUPZTF001")
		ZTF->( dbSeek(xFilial("ZTF")+cTabX+cVerX+cITTX) )
		While !ZTF->(Eof()) .and. ZTF->ZTF_FILIAL+ZTF->ZTF_TABCOM+ZTF->ZTF_VERCOM+ZTF->ZTF_ITTABC == xFilial("ZTF")+cTabX+cVerX+cITTX
			If val(ZTF->ZTF_DIASEM)+1 == _nDia 
				_nVALKM  := ZTF->ZTF_VALKM
				_nVALKMA := ZTF->ZTF_VALKMA
				Exit
			EndIf
			dbSkip()
		EndDo       

		If _nVALKM+_nVALKMA > 0		
		
			For _nX:=1 to Len(oDlgCar:aCols)	
				_nTemp 		:= 0                     
				_nSKMCOM    := oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SKMCOM"})]   
				_nEKMCOM    := oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_EKMCOM"})]
		
				If _nEKMCOM == 0  
					_nTemp := _nSKMCOM *  _nVALKM			
				Else                                                      
					_nTemp1 := _nSKMCOM - _nEKMCOM
					_nTemp2 := _nTemp1 * _nVALKM
					_nTemp3 := _nEKMCOM * _nVALKMA
					_nTemp  := _nTemp3 + _nTemp2
				EndIF

				if ZT1->ZT1_DESLOC == "S"		// Carrega Valor de Deslocamento - Cristiam Rossi em 05/02/2016
					_nTemp4 := ZT1->ZT1_VRDES *  gdFieldGet("ZA7_KMDESL" , _nX, .F., oDlgCar:aHeader, oDlgCar:aCols)
					_nTemp += _nTemp4
			    	gdFieldPut("ZA7_VRDESC" , _nTemp4, _nX, oDlgCar:aHeader, oDlgCar:aCols)
				else
			    	gdFieldPut("ZA7_VRDESC" , 0      , _nX, oDlgCar:aHeader, oDlgCar:aCols)
				endif
				//Pedrassi Begin
				If ZT1->ZT1_TIPLKM == "D" .AND. ZTF->ZTF_VALKM >= 0
					_nTemp	+=	ZT1->ZT1_CUSTO 
				EndIf
				//Pedrassi End
				oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})] := _nTemp
			Next            
			
			
		Else                         
			For _nX:=1 to Len(oDlgCar:aCols)	
				oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})] := 0		
			Next
		EndIF		
		
		// Tratamento especiao para aglutinar o valor total do custo na primeira linha do ZA7
		// Frank Zwarg Fuga 08-03-2016
		// ---------------------------------------------------------------------------------------------------
		_nTotal := 0
		For _nX:=1 to Len(oDlgCar:aCols)
			If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
				_nTotal += oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]
			EndIf
		Next
		_lPrimeira := .T.		    
		For _nX:=1 to Len(oDlgCar:aCols)
			If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
				If _lPrimeira
					oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})] := _nTotal
					_lPrimeira := .F.
				Else
					oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})] := 0
				EndIF
			EndIf
		Next                                                                                         
		// ---------------------------------------------------------------------------------------------------                                                                                                      
		
	Else

		_cMsgAlert := ""
	
		For nX := 1 To Len(oDlgCar:aCols)
	
			cDevEmb := oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_DEVEMB"})]
			cAdic   := oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_ADICIO"})]
			_Munic  := oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_MUNICI"})]
	
			if Select("TZA2") > 0
				TZA2->(dbCloseArea())
			endif
	
			mSQL := " SELECT ZA2_CODIGO,ZA2_DESCRI,ZA2_ESTADO,ZA2_CDPAIS,ZA2_CODMUN, "
			mSQL += " ZT1_CODTAB,ZT1_VERTAB,ZT1_ITEMTB,ZT1_CUSTO,ZT1_TIPVEI,ZT1_PEMERG "
			mSQL += " FROM "+RetSQLName("ZA2")+" ZA2 INNER JOIN "+RetSQLName("ZT1")+" ZT1 "
			mSQL += " ON ZT1_CODORI=ZA2_CODIGO AND ZT1_CODTAB='"+cTabCOM+"' AND ZT1_VERTAB='"+cVerTabC+"'" 
			If	!Empty(cITTabC)
				mSQL += " AND ZT1_ITEMTB='"+cITTabC+"' "
			EndIf
			mSQL += " AND ZT1_CODCLI='"+cCodCLi+"' AND ZT1_LOJCLI='"+cLojCli+"' "
			mSQL += " AND ZT1_TIPTAB='"+cTpTrans+"' AND ZT1_TIPFLU='"+cTpFluxo+"' "
			//mSQL += " AND ZT1_CODORI='"+_Orig+"' AND ZT1_CODDES='"+_Dest+"' "
			mSQL += " WHERE ZT1_FILIAL='"+xFilial("ZT1")+"' AND ZT1.D_E_L_E_T_=' ' "
			mSQL += " AND ZA2_FILIAL='"+xFilial("ZA2")+"' AND ZA2.D_E_L_E_T_=' ' "
			If valtype(dDtIni) == "D"
				mSQL += " AND ZT1_INIVIG <='"+DtoS(dDtIni)+"' AND ( ZT1_FIMVIG >='"+DtoS(dDtIni)+"'  OR ZT1_FIMVIG ='' ) "
			Else                                                                                                          
				mSQL += " AND ZT1_INIVIG <='"+dDtIni+"' AND ( ZT1_FIMVIG >='"+dDtIni+"'  OR ZT1_FIMVIG ='' ) "
			EndIF
		If	Empty(cITTabC)
			mSQL += " AND ZT1_TIPVEI='"+cTpTran+"'"	
		EndIf
		mSQL += " AND ZT1_MSBLQL='2' " //AND ZA2_CODIGO='"+_Munic+"' "
			mSQL += " GROUP BY ZA2_CODIGO,ZA2_DESCRI,ZA2_ESTADO,ZA2_CDPAIS,ZA2_CODMUN, ZT1_CODTAB,ZT1_VERTAB, "
			mSQL += " ZT1_ITEMTB,ZT1_CUSTO,ZT1_TIPVEI, ZT1_PEMERG "
			dbUseArea( .T., "TOPCONN", TCGENQRY(,,mSQL), "TZA2", .F., .T. )
	
			dbSelectArea("TZA2")
			if ! EOF()
				_vFrete := TZA2->ZT1_CUSTO
				nPEMERG := TZA2->ZT1_PEMERG 
				cITTabC := TZA2->ZT1_ITEMTB
				oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_ITTABC"})] := cITTabC
			else
	//			Aviso("A T E N Ç Ã O","O valor informado não consta na Tabela de Custo de Frete "+Trim(Str(nX))+" da aba Coletas" ,{"Ok"}, 2)
				//_cMsgAlert += "O valor informado não consta na Tabela de Compra do Frete "+Trim(Str(nX))+" da aba Coletas" + CRLF
				Loop
			endif
			TZA2->(dbCloseArea())
	        
			cEmerg := oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_EMERGE"})]
	
			if cDevEmb == "R"
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]	:= 0
			elseIf cEmerg =="N" .AND. cDevEmb == "P" .AND. cAdic =="N" // ITEM 1.1.1
				nValor := _vFrete
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]	:= nValor
			Elseif cEmerg =="S" .AND. cDevEmb == "P" .AND. cAdic =="N"  // ITEM 1.1.2
				nValor :=_vFrete+(_vFrete * (nPEMERG/100))
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]	:= nValor
			Elseif cEmerg =="N"  .AND. cDevEmb == "P" .AND. cAdic =="S"	 // ITEM 1.1.3				
				
				IF SELECT("TZTB") > 0
					dbSelectArea("TZTB")
					TZTB->(dbCloseArea())
				Endif
	  
				mSQL := "SELECT ZTB_TABCOM,ZTB_VERTAB,ZTB_CODMUN,ZTB_CUSTO "    
				mSQL += "FROM "+RetSQLName("ZTB")+" ZTB "
				mSQL += " WHERE ZTB_FILIAL ='"+xFilial("ZTB")+"' AND D_E_L_E_T_=' ' "
				mSQL += " AND ZTB_TABCOM='"+cTabCom+"' "
				mSQL += " AND ZTB_VERTAB='"+cVerTabC+"' AND ZTB_CODMUN='"+_Munic+"'"
				mSQL += " AND ZTB_ITTABC='"+cITTabC+"' AND ZTB_MSBLQL != '1' "  
				
				dbUseArea( .T., "TOPCONN", TCGENQRY(,,mSQL), "TZTB", .F., .T. )
	
				dbSelectArea("TZTB")
				IF ! EoF()
					nValor := TZTB->ZTB_CUSTO
					oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]	:= nValor
				else
	//				Aviso("A T E N Ç Ã O","O valor adicional informado não consta na Tabela de Compra do Frete "+Trim(Str(nX))+" da aba Coletas" ,{"Ok"}, 2)
					//_cMsgAlert += "O valor adicional informado não consta na Tabela de Compra do Frete "+Trim(Str(nX))+" da aba Coletas" + CRLF
					oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]	:= 0
					Loop
				Endif
				TZTB->(dbCloseArea())
	
			Elseif cEmerg =="S"  .AND. cDevEmb == "P" .AND. cAdic =="S"	 // ITEM 1.1.4
								
				IF SELECT("TZTB") > 0
					dbSelectArea("TZTB")
					TZTB->(dbCloseArea())
				Endif
	
				mSQL := "SELECT ZTB_TABCOM,ZTB_VERTAB,ZTB_CODMUN,ZTB_CUSTO "    
				mSQL += " FROM "+RetSQLName("ZTB")+" ZTB "
				mSQL += " WHERE ZTB_FILIAL ='"+xFilial("ZTB")+"' AND D_E_L_E_T_=' ' "
				mSQL += " AND ZTB_TABCOM='"+cTabCom+"' "
				mSQL += " AND ZTB_VERTAB='"+cVerTabC+"' AND ZTB_CODMUN='"+_Munic+"'"
				mSQL += " AND ZTB_ITTABC='"+cITTabC+"' AND ZTB_MSBLQL='2' " 						 
	
				dbUseArea( .T., "TOPCONN", TCGENQRY(,,mSQL), "TZTB", .F., .T. )
	
				dbSelectArea("TZTB")
				IF TZTB->(!EoF())
					nValor := TZTB->ZTB_CUSTO +(TZTB->ZTB_CUSTO*nPEMERG/100)
					oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]	:= nValor
				Endif
	
				TZTB->(dbCloseArea())
			Elseif cEmerg =="N" .AND. cDevEmb == "R" .AND. cAdic =="N" // ITEM 1.1.6
				nValor := _vFrete
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]	:= nValor
			Elseif cEmerg =="S"  .AND. cDevEmb == "R" .AND. cAdic =="N"   // ITEM 1.1.7
				if nVEMERG == 0 
					nValor :=_vFrete+(_vFrete * (nPEMERG/100))
				Else
					nValor := nVEMERG
				Endif
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]	:= nValor
	
			Elseif cEmerg =="N"  .AND. cDevEmb == "R" .AND. cAdic =="S"	 // ITEM 1.1.8				
				nValor := _vFrete+(_vFrete * (nPEMERG/100))
				oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]	:= nValor
			Elseif cEmerg =="S"  .AND. cDevEmb == "R" .AND. cAdic =="S"	 // ITEM 1.1.9
	
				IF SELECT("TZTB") > 0
					dbSelectArea("TZTB")
					TZTB->(dbCloseArea())
				Endif
	
				mSQL := "SELECT ZTB_TABCOM,ZTB_VERTAB,ZTB_CODMUN,ZTB_CUSTO "    
				mSQL += " FROM "+RetSQLName("ZTB")+" ZTB "
				mSQL += " WHERE ZTB_FILIAL ='"+xFilial("ZTB")+"' AND D_E_L_E_T_=' ' "
				mSQL += " AND ZTB_TABCOM='"+cTabCom+"' "
				mSQL += " AND ZTB_VERTAB='"+cVerTabC+"' AND ZTB_CODMUN='"+_Munic+"'"
				mSQL += " AND ZTB_ITTABC='"+cITTabC+"' AND ZTB_MSBLQL='2' " 						 
				dbUseArea( .T., "TOPCONN", TCGENQRY(,,mSQL), "TZTB", .F., .T. )
	
				dbSelectArea("TZTB")
				IF TZTB->(!EoF())
					nValor := TZTB->ZTB_CUSTO +(TZTB->ZTB_CUSTO*nPEMERG/100)
					oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]	:= nValor
				Else
	//				Aviso("A T E N Ç Ã O","O valor adicional informado não consta na Tabela de Compra do Frete "+Trim(Str(nX))+" da aba Coletas" ,{"Ok"}, 2)
	//				_cMsgAlert += "O valor adicional informado não consta na Tabela de Compra do Frete "+Trim(Str(nX))+" da aba Coletas" + CRLF
					oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]	:= 0
					Loop
				Endif
	
				TZTB->(dbCloseArea())
			Endif
	
		Next nX
	
	if ! isInCallStack("RECALCVI") .and.  ! Empty( _cMsgAlert ) 	// Cristiam Rossi em 28/01/2016
			Aviso("A T E N Ç Ã O", _cMsgAlert, {"Ok"}, 3)		
		endif
	
	  //	Next nX
				//ZA6_TPTRAN!=!ZT0_TIPTAB!e!ZA6_TIPFLU!=!ZT0_TIPFLU	
	
	EndIF
	

	if valtype(oDlgCar) == "O" .and. !_lGerAcols
		oDlgCar:Refresh()
    	
    	_VTotalV := 0
    	_VTotal  := 0
		For nX := 1 To Len(oDlgCar:aCols)
			_VTotalV += oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_VALOR"})]
			_VTotal  += oDlgCar:aCols[nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_CUSTO"})]
		Next nX    
        
		nValCusTG  := _VTotal
		nValRenTG  := _VTotalV - _VTotal
		nValViagem := _VTotalV 

		oValCusTG:Refresh()
		oValRenTG:Refresh()
 		oValViagem:Refresh()
 			
 		oDlgCar:oBrowse:Refresh()
	Endif	
 	
Return(_Ret)



/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³KMMOD2    ºAutor  ³Frank Zwarg Fuga    º Data ³  02/11/16   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³Encontra o valor do KM para o calculo tipo 4                º±±
±±º          ³_cMod == 1 (Venda), _cMod == 2 (Compra)                     º±±
±±º          ³_cTipRet == 1 retorna o valor, 2 retorna o tipo de processo º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                         º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
User Function KMMOD2(_cMod, _nPonteiro, _cTipRet)
Local _aArea 	:= GetArea()  
Local _cTabV	:= oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TABVEN"})]
Local _cTabC    := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TABCOM"})]
Local _cVerV    := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_VERVEN"})]
Local _cVerC    := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_VERCOM"})]
Local _cIteV    := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_ITTABV"})]
Local _cIteC    := oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_ITTABC"})]
Local _cTpTrans	:= oDlgTra:aCols[oDlgTra:nAt][Ascan(oDlgTra:aHeader,{|x|AllTrim(x[2])=="ZA6_TPTRAN"})]
Local _cCodCli	:= AllTrim(ZA0->ZA0_CLI)
Local _cLojCli	:= AllTrim(ZA0->ZA0_LOJA)
Local _cVenda	:= " "
Local _cCompra  := " "
Local _cRatVen  := " "
Local _cRatCom  := " "
Local _nPerV 	:= 0
Local _nPerC 	:= 0
Local _nRetorno := 0
Local _lErro	:= .F.
Local _nTotal	:= 0
Local _nTemp1   := 0
Local _nTemp2   := 0
Local _nX

// Posicionar na tabela de vendas
ZT0->(dbSetOrder(1))
ZT0->( dbSeek( xFilial("ZT0") + _cTabV + _cVerV + _cCodCLi + _cLojCli + _cTpTrans + _cIteV , .T. ) )
If ZT0->(Eof())
	MsgStop("Não foi possível vincular a tabela de venda.","Atenção!")
	Return .F.
Else
	If ZT0->ZT0_KMPTRE == "S"
		_cVenda := "S" 
		If ZT0->ZT0_RATPES == "S"
			_cRatVen := "S"
		Else               
			_cRatVen := "N"
		EndIF
	Else
		_cVenda := "N"
		If ZT0->ZT0_RATPES == "S"
			_cRatVen := "S"
		Else               
			_cRatVen := "N"
		EndIF
	EndIF              
	_nPerV := ZT0->ZT0_PRTEMB
EndIF                 

// Posicionar na tabela de compras
dbSelectArea("ZT1")
dbSetOrder(1)                                                                                
dbSeek( xFilial("ZT1") + ZT0->ZT0_TABCOM + ZT0->ZT0_VERCOM)
_lAchou := .F.
While !Eof() .and. ZT1_FILIAL+ZT1_CODTAB+ZT1_VERTAB == xFilial("ZT1")+ZT0->ZT0_TABCOM + ZT0->ZT0_VERCOM
	If ZT1_TIPTAB == _cTpTrans .and. ZT1_ITEMTB == ZT0->ZT0_ITTABC .and. ZT1_CODCLI == _cCodCli .and. ZT1_LOJCLI == _cLojCli
		_lAchou := .T.
		Exit
	EndIF
	dbSkip()
EndDo
If Eof()
	_lErro := .T.
Else
	If ZT1->ZT1_KMPTRE == "S"
		_cCompra := "S"
		If ZT0->ZT0_RATPES == "S"
			_cRatCom := "S"
		Else               
			_cRatCom := "N"
		EndIF
	Else
		_cCompra := "N"
		If ZT0->ZT0_RATPES == "S"
			_cRatCom := "S"
		Else               
			_cRatCom := "N"
		EndIF
	EndIF                    
	_nPerC := _nPerV
EndIf	
RestArea(_aArea)                  

If _cTipRet == "2"
	Return {_cVenda,_cCompra,_cRatVen,_cRatCom,_nPerV,_nPerC}
EndIf


If !_lErro
	If _cMod == "1" // Venda
		If _cVenda == "N"
			_nRetorno := oDlgCar:aCols[_nPonteiro][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SKMVEN"})]
		Else
			For _nX:=1 to Len(oDlgCar:aCols)
				If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
					_nTotal += oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SKMVEN"})]
				EndIF
			Next       
			_nRetorno := _nTotal
		EndIf	
	Else
		If _cCompra == "N"
			_nRetorno := oDlgCar:aCols[_nPonteiro][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SKMCOM"})]
		Else     
			For _nX:=1 to Len(oDlgCar:aCols)
				If !oDlgCar:aCols[_nX][Len(oDlgCar:aHeader)+1]
					_nTotal += oDlgCar:aCols[_nX][Ascan(oDlgCar:aHeader,{|x|AllTrim(x[2])=="ZA7_SKMCOM"})]
				EndIF
			Next
			_nRetorno := _nTotal
		EndIf
	EndIF                    
EndIF

Return _nRetorno





********************************************************************************
Static Function fHeader(aCamposSim)
********************************************************************************
Local nPos,aTabAux,aHeader:={}

dbSelectArea("SX3")
dbSetOrder(2)  //X3_CAMPO
For nPos:=1 to Len(aCamposSim)
	If SX3->(dbSeek(PadR(AllTrim(aCamposSim[nPos,1]),Len(X3_CAMPO))))
        aTabAux:={}
		AAdd(aTabAux,TRIM(x3Titulo()))
		AAdd(aTabAux,x3_campo        )
		AAdd(aTabAux,x3_picture      )
		AAdd(aTabAux,x3_tamanho      )
		AAdd(aTabAux,x3_decimal      )
		AAdd(aTabAux,x3_valid        )
		AAdd(aTabAux,x3_usado        )
		AAdd(aTabAux,x3_tipo         )
		AAdd(aTabAux,x3_f3           )
		AAdd(aTabAux,x3_context      )
		AAdd(aTabAux,x3_cbox         )
		AAdd(aTabAux,x3_relacao      )
 		AAdd(aTabAux,x3_when         )

		If !Empty(aCamposSim[nPos,2])
			AAdd(aTabAux,aCamposSim[nPos,2])
		Else
			AAdd(aTabAux,x3_visual       )
		EndIf

		If Empty(AllTrim(x3_vlduser))
			AAdd(aTabAux,"U_ZA0VALID('"+Upper(AllTrim(x3_campo))+"')")
		Else
			AAdd(aTabAux,X3_VLDUSER      )
		EndIf

		AAdd(aTabAux,X3_PICTVAR      )
		AAdd(aTabAux,X3_OBRIGAT      )
		AAdd(aHeader,aTabAux         )
	EndIf
Next

dbSetOrder(1)  //X3_ARQUIVO+X3_ORDEM

Return(AClone(aHeader))

********************************************************************************
Static Function fCols(aHeader,cAlias,nIndice,cChave,cCondicao,cFiltro)
********************************************************************************
Local nPos,aCols0,aCols:={}
Local cAliasAnt:=Alias()

	dbSelectArea(cAlias)

	(cAlias)->(DbSetOrder(nIndice))
	(cAlias)->(DbSeek(cChave,.t.))
	While (cAlias)->(!Eof() .and. &cCondicao)
		If !(cAlias)->(&cFiltro)
			(cAlias)->(DbSkip())
			Loop
		EndIf
		aCols0:={}
		
		For nPos:=1 to Len(aHeader)
			If ! aHeader[nPos,10] == "V"  //x3_context
				(cAlias)->(AAdd(aCols0,FieldGet(FieldPos(aHeader[nPos,2]))))
			Else
				(cAlias)->(AAdd(aCols0,CriaVar(aHeader[nPos,2])))
			EndIf
		Next
		AAdd(aCols0,.F.  )  //Deleted
		AAdd(aCols,aCols0)
		(cAlias)->(DbSkip())
	End

	If Empty(aCols)
		aCols0:={}
		For nPos:=1 to Len(aHeader)
			(cAlias)->(AAdd(aCols0,CriaVar(aHeader[nPos,2])))
		Next
		AAdd(aCols0,.F.  )  //Deleted
		AAdd(aCols,aCols0)
	EndIf

	aCols0:={}
	For nPos:=1 to Len(aHeader)
		(cAlias)->(AAdd(aCols0,CriaVar(aHeader[nPos,2])))
	Next
	AAdd(aCols0,.F.  )  //Deleted

	Do Case
		Case cAlias=="ZA1";oObr_Cols0:={};AAdd(oObr_Cols0,Aclone(aCols0))
		Case cAlias=="ZA4";oRot_Cols0:={};AAdd(oRot_Cols0,Aclone(aCols0))
		Case cAlias=="ZAM";oTre_Cols0:={};AAdd(oTre_Cols0,Aclone(aCols0))
		Case cAlias=="ZA6";oTra_Cols0:={};AAdd(oTra_Cols0,Aclone(aCols0))
		Case cAlias=="ZAG"
			oGru_Cols0:={}
			AAdd(oGru_Cols0,Aclone(aCols0))
		Case cAlias=="ZA5";oGui_Cols0:={};AAdd(oGui_Cols0,Aclone(aCols0))
		Case cAlias=="ZAE";oCon_Cols0:={};AAdd(oCon_Cols0,Aclone(aCols0))
		Case cAlias=="ZA7";oCar_Cols0:={};AAdd(oCar_Cols0,Aclone(aCols0))
		Case cAlias=="ZAA";oRes_Cols0:={};AAdd(oRes_Cols0,Aclone(aCols0))
		Case cAlias=="ZA9";oCus_Cols0:={};AAdd(oCus_Cols0,Aclone(aCols0))
		Case cAlias=="ZA8";oRat_Cols0:={};AAdd(oRat_Cols0,Aclone(aCols0))
		Case cAlias=="ZAI";oDoc_Cols0:={};AAdd(oDoc_Cols0,Aclone(aCols0))
		Case cAlias=="ZAF";oFol_Cols0:={};AAdd(oFol_Cols0,Aclone(aCols0))
		Case cAlias=="ZAD";oLic_Cols0:={};AAdd(oLic_Cols0,Aclone(aCols0))
		Case cAlias=="ZAC";oMao_Cols0:={};AAdd(oMao_Cols0,Aclone(aCols0))
		Case cAlias=="ZA3";oEta_Cols0:={};AAdd(oEta_Cols0,Aclone(aCols0))
		Case cAlias=="ZAK";oAce_Cols0:={};AAdd(oAce_Cols0,Aclone(aCols0))
		Case cAlias=="ZAQ";oAcG_Cols0:={};AAdd(oAcG_Cols0,Aclone(aCols0))
	EndCase

	dbSelectArea(cAliasAnt)

Return(AClone(aCols))










