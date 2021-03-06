#Include "Fivewin.ch"
#Include "xbrowse.ch"
#Include "ado.ch"
#define COMPILAR(x) &(x)
#DEFINE DIRAPP CurDrive()+":\"+CurDir()
#DEFINE ARQAPP GetWinDir()+"\DBFQ.INI"
#DEFINE cTituloAplicacao "DbfQuery Beta � - "+ FWVERSION+" - "+Version()+" - "+str(Year(date()))+" "
REQUEST DBFCDX
Function Main
   Local oRecordSetConsulta:=nil,oBrowserDeDados,oBar,oMsgTagName,oMsgRecorNo
   Public oJanelaPrincipal,oMsgBar,oMsgTime,oIni
   Public cDiretorioDeDados:=cDiretorioDeDados:=+space(300)

   && setar tudo por parametro e remover o maximo de variaveis publicas
   SET STRICTREAD ON
   SET DELETED ON
   SET DATE TO BRITISH
   SET DATE FORMAT "dd/mm/YYYY"
   SET CENTURY ON
   SET EXACT ON
   SET EPOCH TO YEAR(DATE()) - 50
   SET _3DLOOK ON

   SetBalloon( .T. )
   SetResDebug( .t. )
   
   oIni     := TIni():New(ARQAPP)
   if !File(ARQAPP)
      MemoWrit(ARQAPP,"")
      oIni:Set("DBFQUERY","DIRETORIO",DIRAPP)
   endif


   // ADOCONNECT oCn TO MSSQL SERVER 'MATHEUS\SQLEXPRESS' USER 'sa' PASSWORD 'matrix'
   //	oCn:Execute('SET DATEFORMAT YMD') // MUDA DATA PARA O FORMATO USADO PELO SQL
   //	Try
      //		oCn:Execute('CREATE DATABASE INTELIGENCE') // MUDA DATA PARA O FORMATO USADO PELO SQL
   //	End
   //	oCn:Execute('USE INTELIGENCE') // MUDA DATA PARA O FORMATO USADO PELO SQL
   //	SysRefresh()
	
   //	aArquivosDbf:= Directory(DIRAPP+"\*.dbf")
   //	cArquivo:={}
   //	nIni:=time()
   //	for each cArquivo in aArquivosDbf
      //	 	 FW_AdoImportFromDBF( oCn, DIRAPP+"\"+cArquivo[1], StrTran(cArquivo[1],Right(cArquivo[1],4) ))
   //	next
   //	? ElapTime(nIni,time())
	
	
   
   DEFINE ICON   oIconeAplicacao   RESOURCE "ICON"
   DEFINE BITMAP oImagemDeFundo RESOURCE "IMAGEMDEFUNDO"

   DEFINE FONT oFontConsulta NAME "Courier New" Size 7.5,20
   DEFINE FONT oFontBrowser NAME "Courier New" Size 7.5,15
   DEFINE WINDOW oJanelaPrincipal FROM 0, 0 TO 300, 300  ICON oIconeAplicacao  Title cTituloAplicacao
   oJanelaPrincipal:bpainted = {| hdc | palbmpdraw( hdc, 0, 0, oImagemDeFundo:hbitmap,oImagemDeFundo:hPalette,oJanelaPrincipal:nWidth(),oJanelaPrincipal:nHeight())}
   
   @ 0,0 XBROWSE oBrowserDeDados  of oJanelaPrincipal RecordSet oRecordSetConsulta;
    Font oFontBrowser PIXEL AUTOSORT AUTOCOLS CELL LINES ;
    ON CHANGE (mudarMsgBar(oMsgTagName,oMsgRecorNo,oRecordSetConsulta))


   DEFINE BUTTONBAR oBar OF oJanelaPrincipal SIZE 60,75 3DLOOK 2010 
	
   oBrowserDeDados:bClrGrad := oBar:bClrGrad
   oBrowserDeDados:lRecordSelector := .f.

   DEFINE BUTTON oBotaoBar1 OF oBar  MESSAGE "Executar Query" ;
    ACTION ExecutaConsulta(@oRecordSetConsulta,@oBrowserDeDados) PROMPT "Query" PIXEL RESNAME "BTQUERY"

   && MOVIMENTO

   DEFINE BUTTON oBotaoBar1 OF oBar  MESSAGE "Ir para o Primeiro" GROUP RESNAME "BTPRI";
    ACTION (oBrowserDeDados:GoTop(),oBrowserDeDados:Refresh(),Sysrefresh()) PROMPT "Primeiro" PIXEL  WHEN oRecordSetConsulta != nil

   DEFINE BUTTON oBotaoBar2 OF oBar  MESSAGE "Ir para o Anterior" RESNAME "BTANT";
    ACTION (oBrowserDeDados:Skip(-1),oBrowserDeDados:Refresh(),Sysrefresh()) PROMPT "Anterior" PIXEL WHEN oRecordSetConsulta != nil

   DEFINE BUTTON oBotaoBar3 OF oBar  MESSAGE "Ir para o Proximo" RESNAME "BTPRO";
    ACTION (oBrowserDeDados:Skip(),oBrowserDeDados:Refresh(),Sysrefresh()) PROMPT "Pr�ximo" PIXEL  WHEN oRecordSetConsulta != nil

   DEFINE BUTTON oBotaoBar4 OF oBar  MESSAGE "Ir para o Ultimo" RESNAME "BTULT";
    ACTION (oBrowserDeDados:GoBottom(),oBrowserDeDados:Refresh(),Sysrefresh()) PROMPT "�ltimo" PIXEL WHEN oRecordSetConsulta != nil

   && INCLUIR,EDITAR,EXCLUIR
	
   && CREATE INDEX pkclientes ON CLIENTES(CODICLI) with PRIMARY
   DEFINE BUTTON oBotaoBar5 OF oBar  MESSAGE "Incluir Registro Atual" GROUP;
    ACTION RSAppendBlank( @oRecordSetConsulta,@oBrowserDeDados) ,oBrowserDeDados:Refresh(),Sysrefresh() PROMPT "Incluir" PIXEL RESNAME "BTINC" WHEN oRecordSetConsulta != nil

   DEFINE BUTTON oBotaoBar6 OF oBar  MESSAGE "Editar Registro Atual" ;
    ACTION ADOEdit( @oRecordSetConsulta,@oBrowserDeDados ) ,oBrowserDeDados:Refresh(),Sysrefresh()  PROMPT "Alterar" PIXEL RESNAME "BTALT" WHEN oRecordSetConsulta != nil

   DEFINE BUTTON oBotaoBar7 OF oBar  MESSAGE "Excluir Registro Atual" ;
    ACTION RSDelRecord( @oRecordSetConsulta,@oBrowserDeDados ) PROMPT "Excluir" PIXEL RESNAME "BTEXC" WHEN oRecordSetConsulta != nil

   && Filtro

   DEFINE BUTTON oBotaoBar8 OF oBar  MESSAGE "Filtro na Tabela" GROUP;
    ACTION pesquisanoResultSet(@oRecordSetConsulta,@oBrowserDeDados) PROMPT "Filtro" PIXEL RESNAME "BTPESQ" WHEN oRecordSetConsulta != nil

   DEFINE BUTTON oBotaoBar8_1 OF oBar  MESSAGE "Indice Primario da Tabela" ;
    ACTION DefineChavePrimaria(@oRecordSetConsulta,@oBrowserDeDados) PROMPT "Indice" PIXEL RESNAME "BTIND" WHEN oRecordSetConsulta != nil .and. oRecordSetConsulta:Fields:Count >= 40

   && Utilitarios

   DEFINE BUTTON oBotaoBar9 OF oBar  MESSAGE "Imprimir Dados" RESNAME "BTPRINT";
    ACTION (MsgRun("Gerando Processo ...","Aguarde",{||oBrowserDeDados:Report()}),oBrowserDeDados:Refresh(),Sysrefresh()) PROMPT "Imprimir" PIXEL GROUP WHEN oRecordSetConsulta != nil

   DEFINE BUTTON oBotaoBar10 OF oBar  MESSAGE "Exportar dados para Excel" RESNAME "BTEXCEL";
    ACTION (MsgRun("Gerando Processo ...","Aguarde",{||oBrowserDeDados:ToExcel()}),oBrowserDeDados:Refresh(),Sysrefresh()) PROMPT "P/Excel" PIXEL WHEN oRecordSetConsulta != nil

   DEFINE BUTTON oBotaoBar11 OF oBar  MESSAGE "Exportar dados para dbf" RESNAME "BTDBF";
    ACTION (MsgRun("Gerando Processo ...","Aguarde",{||oBrowserDeDados:ToDbf(DIRAPP+"\Query.dbf")}),iif (file(DIRAPP+"\Query.dbf"),Msginfo("Arquivo Criado com Sucesso em "+DIRAPP+"\Query.dbf","Informa��o"),MsgAlert("Falha ao criar "+DIRAPP+"\Query.dbf","Informa��o")),oBrowserDeDados:Refresh(),Sysrefresh()) PROMPT "P/Dbf" PIXEL WHEN oRecordSetConsulta != nil

   && Apoio
	
   DEFINE BUTTON oBotaoBar12 OF oBar  MESSAGE "Configurar Diretorio" GROUP;
    ACTION configuraDiretorio() PROMPT "Config." PIXEL RESNAME "BTDIRE"

   DEFINE BUTTON oBotaoBar13 OF oBar  MESSAGE "Sair do DbfQuery" ;
    ACTION oJanelaPrincipal:End() PROMPT "Sair" PIXEL  RESNAME "BTFECHAR"



   DEFINE MSGBAR oMsgBar OF oJanelaPrincipal PROMPT "www.matheusfarias.com.br" 2010
   DEFINE MSGITEM oMsgTime OF oMsgBar PROMPT '' SIZE 250
   DEFINE MSGITEM oMsgRecorNo OF oMsgBar PROMPT '' SIZE 150
   DEFINE MSGITEM oMsgTagName OF oMsgBar PROMPT '' SIZE 150
   oMsgBar:DateOn()
   oMsgBar:ClockOn()
   oMsgBar:KeybOn()


   ACTIVATE WINDOW oJanelaPrincipal maximized
Function mudarMsgBar(oMsgTagName,oMsgRecorNo,oRecordSetConsulta)
    oMsgTagName:Settext("Ordenado" + ": " + Iif( Empty( oRecordSetConsulta:Sort ),FWString( "natural order" ), oRecordSetConsulta:Sort))
    oMsgRecorNo:Settext("Reg.Atual: " + AllTrim( Str( If( oRecordSetConsulta:AbsolutePosition == -3, oRecordSetConsulta:RecordCount() + 1,oRecordSetConsulta:AbsolutePosition ) ) ) + " / " + AllTrim( Str( oRecordSetConsulta:RecordCount() ) ))
    oRecordSetConsulta:Update()
    
Function pesquisanoResultSet(oRecordSetConsulta,oBrowserDeDados)
   Local oDlgPesquisa,oBrush
   Local oComboCampos,cComboCampos:="",aComboCampos:={}
   Local oGetPesquisa,cGetPesquisa:=space(200)
   Local lsave:=.f.,n:=0,oRadioFiltro,nOpcaoFiltro:=1

   for n = 1 to oRecordSetConsulta:Fields:Count
      AAdd(aComboCampos,oRecordSetConsulta:Fields[ n - 1 ]:Name)
   next

   Define Brush oBrush Resource "PEDRA"
   Define Dialog oDlgPesquisa Title "DbfQuery Beta � - Pesquisa" RESOURCE "DLG_PESQUISA" BRUSH oBrush Transparent
			
   REDEFINE COMBOBOX oComboCampos var cComboCampos ITEMS aComboCampos id 4002 of oDlgPesquisa
   REDEFINE GET      oGetPesquisa var cGetPesquisa id 4004 of oDlgPesquisa
   REDEFINE RADIO    oRadioFiltro var nOpcaoFiltro id 4008,4009 of oDlgPesquisa
			
   redefine buttonbmp obotaook id 4006 of oDlgPesquisa BITMAP "BTOK" ADJUST ACTION (lsave:=.t.,oDlgPesquisa:End())
   redefine buttonbmp obotaoca id 4007 of oDlgPesquisa BITMAP "BTCA" ADJUST ACTION (lsave:=.f.,oDlgPesquisa:End())

   Activate Dialog oDlgPesquisa
			
   if lsave
      oBrowserDeDados:lIncrFilter   := (nOpcaoFiltro == 1)
      oBrowserDeDados:lSeekWild     := (nOpcaoFiltro == 2)
      oBrowserDeDados:cFilterFld    := alltrim(cComboCampos)
      oBrowserDeDados:nStretchCol   := STRETCHCOL_WIDEST
      oBrowserDeDados:Seek(alltrim(cGetPesquisa))
   endif

Function ExecutaConsulta(oRecordSetConsulta,oBrowserDeDados) && colocar parametro
   Local oDlgConsulta
   Local cSqlSintax:=oIni:Get("DBFQUERY","CONSULTA","")+space(5000),oReturn:=nil
   Local lsave:=.F.
         
   Define Brush oBrush Resource "PEDRA"
   Define Dialog oDlgConsulta Title "DbfQuery Beta � - Executa Sql" Resource "DLG_CONSULTA" BRUSH oBrush Transparent

   REDEFINE GET oGetConsulta var cSqlSintax id 4001 of oDlgConsulta TEXT Font oFontConsulta
			
   redefine buttonbmp obotaook id 4003 of oDlgConsulta BITMAP "BTOK" ADJUST ACTION (lsave:=.t.,oDlgConsulta:End())
   redefine buttonbmp obotaoca id 4002 of oDlgConsulta BITMAP "BTCA" ADJUST ACTION (lsave:=.f.,oDlgConsulta:End())

   Activate Dialog oDlgConsulta
   IF LSAVE
      if Empty(cSqlSintax)
         return .f.
      else
         oIni:Set("DBFQUERY","CONSULTA",cSqlSintax)
         hini:=time()
         MsgRun("Processando Consulta dos dados ....","Aguarde",{||oReturn:=SqlQuery( cSqlSintax )})
         if oReturn != nil .and. oReturn:cClassname == "ADODB.RecordSet"
            oRecordSetConsulta:=oReturn
            oBrowserDeDados:SetAdo( oRecordSetConsulta, .t., .t.)
            oBrowserDeDados:CreateFromCode()
            oJanelaPrincipal:oClient  := oBrowserDeDados
         else
            if oRecordSetConsulta != nil
               oRecordSetConsulta:Update()
            endif
         endif
         oMsgTime:Settext("Tempo da Ultima Consulta: "+ ElapTime(hini,time()))
      ENDIF
      oMsgBar:Refresh()
      oBrowserDeDados:Refresh()
      oJanelaPrincipal:Resize()
      oJanelaPrincipal:Refresh()
      Sysrefresh()
   endif
Function configuraDiretorio()
   Local oDlgDiretorio,oBrush,lsave:=.f.
   Local oGetDiretorioDeDados,cDiretorioDeDados:=""

   cDiretorioDeDados:=oIni:Get("DBFQUERY","DIRETORIO",DIRAPP)+space(300)

   Define Brush oBrush Resource "PEDRA"
   Define Dialog oDlgDiretorio Title "DbfQuery Beta � - Configurar Diretorio" Resource "JANELA_DIRETORIO" BRUSH oBrush Transparent

   REDEFINE GET oGetDiretorioDeDados var cDiretorioDeDados id 4002 of oDlgDiretorio ACTION (cDiretorioDeDados:=cGetDir32("Informe a Pasta","c:\"),oGetDiretorioDeDados:Refresh()) BITMAP "BTFOLDER"
   redefine buttonbmp obotaook id 4003 of oDlgDiretorio BITMAP "BTOK" ADJUST ACTION (lsave:=.t.,oDlgDiretorio:End())
   redefine buttonbmp obotaoca id 4004 of oDlgDiretorio BITMAP "BTCA" ADJUST ACTION (lsave:=.f.,oDlgDiretorio:End())

   Activate Dialog oDlgDiretorio

   if lsave
      if !Empty(cDiretorioDeDados)
         oIni:Set("DBFQUERY","DIRETORIO",cDiretorioDeDados)
      endif
   endif

Function SqlQuery( cSqlSintax )
   local uRet    := nil , lExecute:=.f.
   Local oSqlConexao := FW_OpenAdoConnection( [Provider=Microsoft.Jet.OLEDB.4.0;Data Source=]+oIni:Get("DBFQUERY","DIRETORIO",DIRAPP)+[ ;Extended Properties=dBASE IV;User ID=Admin;Password=] ,.t.)
   
    cSql     := Upper( alltrim(cSqlSintax) )

   lExecute  := !( LEFT( cSql, 7 ) == "SELECT " )

   if oSqlConexao != nil
      if lExecute
         TRY
            uRet := oSqlConexao:Execute( cSql )
         catch oError
            uRet:=nil
            MsgAlert("Erro Tecnico:"+oError:Description,"Alerta")
         END
      else
         Try
            uRet      := FW_OpenRecordSet( oSqlConexao, cSql )
         Catch oError
            uRet:=nil
            MsgAlert("Erro Tecnico:"+oError:Description,"Alerta")
         End
      endif
   endif
   return uRet
Function DefineChavePrimaria(oRecordSetConsulta,oBrowserDeDados)
   Local oDlgChave,oLbxCampos,oLbxIndice,cCampo:="",cIndice:="",aLbxCampos:={},aLbxIndice:={},lsave:=.f.,cTabela:="",oComboTav,aTabelas:={}
   aTabelas:= sqltabelas()
	for n = 1 to oRecordSetConsulta:Fields:Count
      AAdd(aLbxCampos,oRecordSetConsulta:Fields[ n - 1 ]:Name)
   next
   Define Brush oBrush Resource "PEDRA"
   Define Dialog oDlgChave Resource "DLG_CHAVEPRIMARIA" Brush oBrush Transparent Title "DbfQuery Beta � - Chave prim�ria"

   REDEFINE LISTBOX oLbxCampos var cCampo  ITEMS aLbxCampos ID 4001 OF oDlgChave
   REDEFINE LISTBOX oLbxIndice var cIndice ITEMS aLbxIndice ID 4002 OF oDlgChave

   REDEFINE COMBOBOX oComboTav var cTabela ITEMS aTabelas id 4010 OF oDlgChave valid ( ! Empty(cTabela))

   REDEFINE BUTTON ID 4003 OF oDlgChave ACTION (iif (Ascan(aLbxIndice,cCampo) == 0,AAdd(aLbxIndice,cCampo), MsgAlert("Campo j� adicionado!","Aten��o")),oLbxIndice:SetItems( aLbxIndice ),oLbxCampos:Refresh(),oLbxIndice:Refresh(),oLbxIndice:Update(),sysrefresh())
   REDEFINE BUTTON ID 4004 OF oDlgChave ACTION ( ADel(aLbxIndice,Ascan(aLbxIndice,cIndice),.t. ),oLbxCampos:Refresh(),oLbxIndice:SetItems( aLbxIndice ),oLbxIndice:Refresh(),oLbxIndice:Update(),sysrefresh()) when (len(aLbxIndice)>0)
	      
   redefine buttonbmp obotaook id 4005 of oDlgChave BITMAP "BTOK" ADJUST ACTION (lsave:=.t.,oDlgChave:End())
   redefine buttonbmp obotaoca id 4006 of oDlgChave BITMAP "BTCA" ADJUST ACTION (lsave:=.f.,oDlgChave:End())

   ACTIVATE DIALOG oDlgChave
	      
   if lsave
      cSqlSintaxIndex:='CREATE INDEX pk'+cTabela+' ON '+cTabela+'('+StrTran(arraytotext(aLbxIndice),'"')+') with PRIMARY'
      MsgRun("Atualizando Informa��es ... ","Aguarde ...",{||SqlQuery(cSqlSintaxIndex)})
      oRecordSetConsulta:Update()

   endif

static function RSAppendBlank( oRs )

   local n, aValues := {}

   if oRs:RecordCount() > 0
      oRs:MoveLast()

      for n = 1 to oRs:Fields:Count
         AAdd( aValues, oRs:Fields[ n - 1 ]:Value )
         if n == 1 .and. ValType( aValues[ 1 ] ) == "N"
            aValues[ 1 ]++
         else
            aValues[ n ] = uValBlank( aValues[ n ] )
            if ValType( aValues[ n ] ) == "D" .and. Empty( aValues[ n ] )
               aValues[ n ] = Date()
            endif
         endif
      next
   else
      aValues = Array( oRs:Fields:Count )
      for n = 1 to oRs:Fields:Count
         do case
         case oRs:Fields[ n - 1 ]:Type == 3
            aValues[ n ] = 1

         case oRs:Fields[ n - 1 ]:Type == 202 .or. ;
             oRs:Fields[ n - 1 ]:Type == 203
            aValues[ n ] = Space( oRs:Fields[ n - 1 ]:DefinedSize )

         case oRs:Fields[ n - 1 ]:Type == 131 .or. oRs:Fields[ n - 1 ]:Type == 16 .or. ;
             oRs:Fields[ n - 1 ]:Type == 2 .or. oRs:Fields[ n - 1 ]:Type == 11
            aValues[ n ] = 0

         case oRs:Fields[ n - 1 ]:Type == 135 .or. oRs:Fields[ n - 1 ]:Type == 7
            aValues[ n ] = Date()

         otherwise
            MsgInfo( oRs:Fields[ n - 1 ]:Type )
         endcase
      next
   endif

   oRs:AddNew()

   if ! Empty( aValues )
      for n = 1 to oRs:Fields:Count
         try
            oRs:Fields[ n - 1 ]:Value = aValues[ n ]
         end
      next
   endif

   oRs:Update()

   return nil
static function RSDelRecord( oRecordSetConsulta )
   local n := oRecordSetConsulta:AbsolutePosition
   if ! MsgYesNo( "Deseja Excluir Esse Registro ?" ,"Pergunta")
      return nil
   endif
   Try
      oRecordSetConsulta:Delete()
      oRecordSetConsulta:Update()
      if n > oRecordSetConsulta:RecordCount()  // Happens only when last record is deleted
         n--
      endif
      oRecordSetConsulta:AbsolutePosition := n  // in most cases this n is not changed. but this assignment is necessary.
      SysRefresh()
   Catch oError
      MsgAlert("N�o foi possivel excluir o registro escolhido por favor, acesse o botao Indice e configure esta informa��o. "+CRLF+"Erro Tecnico:"+oError:Description)
   End

   return nil

   //----------------------------------------------------------------------------//
static function RSLoadRecord( oRecordSet )

   local aRecord := {}, n

   for n = 1 to oRecordSet:Fields:Count
      AAdd( aRecord, { oRecordSet:Fields[ n - 1 ]:Name, oRecordSet:Fields[ n - 1 ]:Value } )
      If ValType( ATail( aRecord )[ 2 ] ) == "C"
         ATail( aRecord )[ 2 ] = PadR( ATail( aRecord )[ 2 ], Min( oRecordSet:Fields[ n - 1 ]:DefinedSize, 255 ) )
      endif
   next

   return aRecord
function SetEditType( oRs, oBrw, oBtnSave )

   local cType, cAlias

   if Empty( oRs )
      cType  = FieldType( oBrw:nArrayAt )
      cAlias = Alias()
   else
      cType = FWAdoFieldType( oRs, oBrw:nArrayAt )
   endif

   do case
   case cType == "M"
      oBrw:aCols[ 2 ]:nEditType = EDIT_BUTTON
      if Empty( oRs )
         oBrw:aCols[ 2 ]:bEditBlock = ;
          { || If( ( cAlias )->( EditMemo( oBrw ) ), oBtnSave:Enable(),) }
      else
         oBrw:aCols[ 2 ]:bEditBlock = ;
          { || If( EditMemo( oBrw ), oBtnSave:Enable(),) }
      endif

   case cType == "D"
      oBrw:aCols[ 2 ]:nEditType = EDIT_BUTTON
      oBrw:aCols[ 2 ]:bEditBlock = { || If( ! Empty( oBrw:aRow[ 2 ] ) .and. ;
       ! AllTrim( DtoC( oBrw:aRow[ 2 ] ) ) == "/  /",;
       MsgDate( oBrw:aRow[ 2 ] ),;
       MsgDate( Date() ) ) }

   case cType == "L"
      oBrw:aCols[ 2 ]:nEditType = EDIT_LISTBOX
      oBrw:aCols[ 2 ]:aEditListTxt   = { ".T.", ".F." }
      oBrw:aCols[ 2 ]:aEditListBound = { .T., .F. }

   otherwise
      oBrw:aCols[ 2 ]:nEditType = EDIT_GET
   endcase

   return nil
function ADOEdit( oRecordSetConsulta )

   local oWnd, aRecord, oBar, oBrw, oMsgBar
   local oBtnSave, nRecNo := oRecordSetConsulta:BookMark
   local oMsgDeleted

   aRecord = RSLoadRecord( oRecordSetConsulta )

   DEFINE WINDOW oWnd TITLE "DbfQuery Beta � - Editar"

   DEFINE BUTTONBAR oBar OF oWnd 2010 SIZE 70, 70

   DEFINE BUTTON oBtnSave OF oBar PROMPT "Salvar" RESOURCE "BTSAVE" ;
    ACTION ( FWAdoSaveRecord( oRecordSetConsulta, aRecord, nRecNo ) , oBtnSave:Disable(), oBrw:SetFocus() )

   oBtnSave:Disable()

   DEFINE BUTTON OF oBar PROMPT "Anterior" RESOURCE "BTANT" ;
    ACTION ( If( oRecordSetConsulta:AbsolutePosition > 1,;
    ( oRecordSetConsulta:MovePrevious(),;
    nRecNo := oRecordSetConsulta:BookMark,;
    oBrw:SetArray( RSLoadRecord( oRecordSetConsulta ) ),;
    oBrw:Refresh(), Eval( oBrw:bChange ) ),), oBrw:SetFocus() ) GROUP

   DEFINE BUTTON OF oBar PROMPT "Pr�ximo" RESOURCE "BTPRO" ;
    ACTION ( If( oRecordSetConsulta:AbsolutePosition < oRecordSetConsulta:RecordCount,;
    ( oRecordSetConsulta:MoveNext(),;
    nRecNo := oRecordSetConsulta:BookMark,;
    oBrw:SetArray( RSLoadRecord( oRecordSetConsulta ) ),;
    oBrw:Refresh(), Eval( oBrw:bChange ) ),), oBrw:SetFocus() )

   DEFINE BUTTON OF oBar PROMPT "Imprimir" RESOURCE "BTPRINT" ;
    ACTION oBrw:Report() GROUP

   DEFINE BUTTON OF oBar PROMPT "Ver" RESOURCE "BTSHOW" ;
    ACTION View( oBrw:aRow[ 2 ], oWnd )

   DEFINE BUTTON OF oBar PROMPT "Fechar" RESOURCE "BTRET" ;
    ACTION oWnd:End() GROUP

   @ 0, 0 XBROWSE oBrw OF oWnd ARRAY aRecord AUTOCOLS LINES ;
    HEADERS "Campo", "Valor" COLSIZES 150, 400 FASTEDIT ;
    ON CHANGE ( SetEditType( oRecordSetConsulta, oBrw, oBtnSave ), oBrw:DrawLine( .T. ),;
    oMsgBar:cMsgDef := " RecNo: " + AllTrim( Str( oRecordSetConsulta:AbsolutePosition ) ) + ;
    "/" + AllTrim( Str( oRecordSetConsulta:RecordCount ) ),;
    oMsgBar:Refresh() )

   oBrw:nEditTypes = EDIT_GET
   oBrw:aCols[ 1 ]:nEditType = 0 // Don't allow to edit first column
   oBrw:aCols[ 2 ]:bOnChange = { || oBtnSave:Enable() }
   oBrw:aCols[ 2 ]:lWillShowABtn = .T.
   oBrw:nMarqueeStyle = MARQSTYLE_HIGHLROW
   oBrw:bClrSel = { || { CLR_WHITE, RGB( 0x33, 0x66, 0xCC ) } }
   oBrw:SetColor( CLR_BLACK, RGB( 232, 255, 232 ) )
   oBrw:CreateFromCode()
   oBrw:SetFocus()

   oWnd:oClient = oBrw

   DEFINE MSGBAR oMsgBar ;
    PROMPT " RecNo: " + AllTrim( Str( oRecordSetConsulta:AbsolutePosition ) ) + "/" + ;
    AllTrim( Str( oRecordSetConsulta:RecordCount ) ) OF oWnd 2010

   ACTIVATE WINDOW oWnd

   return nil

function FWAdoSaveRecord( oRS, aRecord, nRecNo )

   local n, oField, uVal, uNew
   local lUpdated := .f., lSaved   := .f.

   if ! Empty( nRecNo ) .and. oRs:BookMark != nRecNo
      oRs:BookMark = nRecNo
   endif

   for n = 1 to oRS:Fields:Count
      oField = oRs:Fields( n - 1 )
      if FW_AdoFieldUpdateable( oRs, oField ) == .f.
         LOOP
      endif
      uVal   = oField:Value
      uNew   = aRecord[ n, 2 ]
      if Empty( uVal ) .and. Empty( uNew )
         LOOP
      endif

      #ifdef __XHARBOUR__
         if Empty( uNew ) .and. lAnd( oField:Attributes, 0x20 ) // nullable field
            oField:Value = VTWrapper( 1 )  // assigning NULL
            LOOP
         endif
      #endif
      // assume that uNew is not NIL .and. is correct data type
      if ValType( uNew ) == 'C'
         if oField:Type == adChar // Fixed width
            uNew = PadR( uNew, oField:DefinedSize )
         else
            uNew = Left( Trim( uNew ), oField:DefinedSize )
         endif
      endif

      if ! ( ValType( uVal ) == ValType( uNew ) .and. uVal == uNew )
         if AScan( { adBinary, adVarBinary, adLongVarBinary }, oField:Type ) != 0
            uNew = HB_StrToHex( uNew )
         endif

         #ifndef __XHARBOUR__
            // Harbour has problem in assigning Empty Dates
            //
            if ValType( uVal ) $ 'DT' .and. Empty( uNew ) .and. ;
                ! ( FW_RDBMSName( oRs:ActiveConnection ) == "MSACCESS" )
               uNew = 0
            endif
         #endif

         /*
         if ValType( uNew ) == "L"
            uNew        = If( uNew, 1, 0 )
         endif
         */

         TRY
            oField:Value = uNew
            lUpdated     = .T.
         catch oError
            ? oField:Name, uNew
         END
      endif
   next

   if lUpdated
      TRY
         oRS:Update()
         lSaved   := .t.
      catch oError
         MsgAlert("Aten��o pode ser necessario criar uma chave primaria para sua tabela, acesse o botao Indice e configure esta informa��o. "+CRLF+"Erro Tecnico:"+oError:Description,"Alerta")
         oRS:CancelUpdate()
      END
   endif

   return lSaved

function EditMemo( oBrw )

   local cTemp  := oBrw:aRow[ 2 ]
   local lResult := .F.

   if lResult := MemoEdit( @cTemp, oBrw:aRow[ 1 ] )
      oBrw:aRow[ 2 ] = cTemp
      oBrw:DrawLine()
   endif

   return lResult
function View( cFileName, oWnd )

   local cExt

   if ! File( cFileName )
      return nil
   endif

   cExt = Lower( cFileExt( cFileName ) )

   do case
   case cExt == "bmp"
      WinExec( "mspaint" + " " + cFileName )

   case cExt == "txt"
      WinExec( "notepad" + " " + cFileName )

   otherwise
      ShellExecute( oWnd:hWnd, "Open", cFileName )
   endcase

   return nil
Function sqltabelas()
   Local oSqlConexao,oResultSetTabelas:=nil
   Local aTabelas:={}
   oSqlConexao := FW_OpenAdoConnection( [Provider=Microsoft.Jet.OLEDB.4.0;Data Source=]+oIni:Get("DBFQUERY","DIRETORIO",DIRAPP)+[ ;Extended Properties=dBASE IV;User ID=Admin;Password=] ,.t.)
	if oSqlConexao!=nil
		oResultSetTabelas = oSqlConexao:OpenSchema( 20 ) // adSchemaTables
	   oResultSetTabelas:Filter = "TABLE_TYPE='TABLE'"
	   oResultSetTabelas:MoveFirst()
	   While !oResultSetTabelas:Eof()
	      AAdd(aTabelas,oResultSetTabelas:Fields("TABLE_NAME"):Value)
	      oResultSetTabelas:MoveNext()
	   end
	endif   
   return aTabelas
function curdrive()
   return hb_curdrive()


