'HASH: D3354E42FB1182BF10E9AF3625610F91
Option Explicit

'#Uses "*Modulo11"

Dim x As Object
Dim GPARAR As Boolean

Dim viBanco As Integer
Dim viAgencia As Integer
Dim VCC As String
Dim vDV As String
Dim Msg As String

Public Function IsInt(pValor As String) As Boolean
  Dim vAux As Long

  On Error GoTo Erro
  vAux = CLng(pValor)
  IsInt = True
   

  Exit Function
Erro:
  IsInt = False

End Function



Public Sub APAGARPF_OnClick()


Dim teste As Long

NewCounter2("NUMERO_SEI", 0, 1, teste)
MsgBox(Str(teste))





Exit Sub

  Dim dllValidarMensagem As Object
  Set dllValidarMensagem = CreateBennerObject("Benner.Saude.WSTiss.Versionador.VersionadorImportarMensagemTISS")
  SessionVar("HANDLE") = CStr(CurrentQuery.FieldByName("TEXTO").AsString)
  SessionVar("HANDLETABELA_TISS") = CStr(CurrentQuery.FieldByName("TEXTO").AsString)
  SessionVar("NOMETABELA_TISS") = "SAM_PRESTADOR_MENSAGEMTISS"
  SessionVar("NOMECAMPO_TISS") = "ARQUIVORECEBIDO"
  SessionVar("HANDLE_TISVERSAO") = "0"
  dllValidarMensagem.Exec(CurrentSystem)
  Set dllValidarMensagem = Nothing
  Exit Sub

  Dim qdoc As Object
  Dim qfat As Object
  Dim qupfat As Object
  Dim qupdoc As Object
  Dim q As Object
  Dim sqlwhere As String
  Dim opBaixa As Long
  Dim lancs As String
  Dim dll As Object
  Dim verro As Integer
  Dim CONTADOR As Integer







  Set qdoc = NewQuery
  Set qfat = NewQuery
  Set qupfat = NewQuery
  Set qupdoc = NewQuery
  Set q = NewQuery
  Set dll = CreateBennerObject("SFNBAIXA.DOCUMENTO")

  sqlwhere = "CANCDATA IS NULL AND DATAVENCIMENTO >'20030501'"


  If sqlwhere = "" Then
    MsgBox("defina a condicção de seleçào na macro")
    Exit Sub
  End If
  qdoc.Clear
  qdoc.Add("SELECT COUNT(*) NREC FROM SFN_DOCUMENTO D WHERE  " + sqlwhere)
  qdoc.Active = True
  If MsgBox("registros a processar: " + qdoc.FieldByName("NREC").AsString + " . Deseja continuar?", vbYesNo) = vbNo Then
    qdoc.Active = False
    GoTo finaliza
  End If
  CONTADOR = qdoc.FieldByName("NREC").AsInteger

  qdoc.Clear
  qdoc.Add("SELECT D.HANDLE,D.VALORTOTAL, D.DATAVENCIMENTO, D.NUMERO, D.TESOURARIA FROM SFN_DOCUMENTO D WHERE  " + sqlwhere)
  qdoc.Active = True
  q.Clear
  q.Add("SELECT HANDLE FROM SIS_OPERACAO WHERE CODIGO=130")'OPERACAOBAIXA
  q.Active = True
  opBaixa = q.FieldByName("HANDLE").AsInteger
  q.Active = False

  qupfat.Add("UPDATE SFN_FATURA SET SALDO=VALOR, BAIXADATA=NULL, BAIXAVALOR=NULL, BAIXAJURO=NULL, BAIXAMULTA=NULL, BAIXACORRECAO=NULL, BAIXADESCONTO=NULL WHERE ")
  qupfat.Add("HANDLE IN (SELECT DF.FATURA FROM SFN_DOCUMENTO_FATURA DF WHERE DF.DOCUMENTO=:DOCUMENTO)")

  qupdoc.Add("UPDATE SFN_DOCUMENTO SET BAIXAVALOR=NULL, BAIXADATA=NULL, BAIXAJURO=NULL, BAIXAMULTA=NULL, BAIXACORRECAO=NULL, BAIXADESCONTO=NULL, BAIXAMOTIVO=NULL WHERE HANDLE=:DOCUMENTO")
  GPARAR = False
  While(Not qdoc.EOF)And(Not GPARAR)
  'excluir os lançamentos de baixa das faturas ligadas a este documento
  StartTransaction
  On Error GoTo rolbeca

  lancs = "SELECT L.HANDLE FROM SFN_FATURA_LANC L, SFN_DOCUMENTO_FATURA DF WHERE DF.DOCUMENTO=" + qdoc.FieldByName("HANDLE").AsString + " AND L.FATURA=DF.FATURA AND L.OPERACAO=" + Str(opBaixa)
  q.Clear
  q.Add("DELETE FROM SFN_CONTAB_LANC_DEBCRE WHERE HANDLE IN(SELECT E.HANDLE ")
  q.Add("  FROM SFN_CONTAB_LANC_DEBCRE E, SFN_CONTAB_LANC D")
  q.Add(" WHERE E.CONTABLANC = D.HANDLE")
  q.Add("   AND D.FATURALANC  IN (" + lancs + "))")
  q.ExecSQL

  q.Clear
  q.Add("DELETE FROM SFN_CONTAB_LANC WHERE HANDLE IN(SELECT D.HANDLE ")
  q.Add("  FROM SFN_CONTAB_LANC D")
  q.Add(" WHERE D.FATURALANC  IN (" + lancs + "))")
  q.ExecSQL

  q.Clear
  q.Add("DELETE FROM SFN_FATURA_LANC_CC WHERE HANDLE IN(SELECT D.HANDLE ")
  q.Add("  FROM SFN_FATURA_LANC_CC D")
  q.Add(" WHERE D.LANCAMENTO  IN (" + lancs + "))")
  q.ExecSQL

  q.Clear
  q.Add("DELETE FROM SFN_FATURA_LANC WHERE HANDLE IN (" + lancs + ")")
  q.ExecSQL

  'atualizar a fatura
  qupfat.ParamByName("DOCUMENTO").AsInteger = qdoc.FieldByName("HANDLE").AsInteger
  qupfat.ExecSQL

  'atualizar o documento
  qupdoc.ParamByName("DOCUMENTO").AsInteger = qdoc.FieldByName("HANDLE").AsInteger
  qupdoc.ExecSQL

  'baixar o documento
  verro = dll.BxDocOnLine(CurrentSystem, qdoc.FieldByName("HANDLE").AsInteger, _
          qdoc.FieldByName("DATAVENCIMENTO").AsDateTime, _
          qdoc.FieldByName("DATAVENCIMENTO").AsDateTime, _
          qdoc.FieldByName("TESOURARIA").AsInteger, _
          qdoc.FieldByName("VALORTOTAL").AsFloat, _
          0, _
          0, _
          0, _
          0, _
          True, _
          0, _
          0, _
          0)


  '      0,'JURO
  '      0, _ 'MULTA
  '      0, _ 'CORRECAO
  '      0, _ 'DESCONTO
  '      True, _ 'CHECKEXCECAO
  '      0, _ 'TESOURARIALANC
  '      0, _ 'ROTARQUIVO
  '      0)'ROTDOCUMENTO


  If verro <0 Then
    MsgBox("documento " + qdoc.FieldByName("NUMERO").AsString + " não pode ser baixado pelo erro:" + Str(verro))
    Rollback
  Else
    Commit
  End If

  GoTo prox

rolbeca :
  MsgBox("documento " + qdoc.FieldByName("NUMERO").AsString + " não pode ser baixado por erro no processo")
  Rollback

prox :

  qdoc.Next
  CONTADOR = CONTADOR -1
  ABREVIATURA.Caption = Str(CONTADOR)

Wend

If GPARAR = True Then
  MsgBox("INTERROMPIDO PELO USUARIO")
Else
  MsgBox("FINALIZADO")
End If
finaliza :
Set qdoc = Nothing
Set qfat = Nothing
Set qupfat = Nothing
Set qupdoc = Nothing
Set q = Nothing
End Sub

Public Sub gravaAbrev( ptexto As String)
	Dim sql As Object
	Set sql=  NewQuery

	On Error GoTo except
		sql.Add("INSERT INTO ABREV(HANDLE, TEXTO) VALUES (:HANDLE, :TEXTO)")
		sql.ParamByName("HANDLE").AsInteger = NewHandle("ABREV")
		sql.ParamByName("TEXTO").AsString = ptexto
		sql.ExecSQL
	except:
		Set sql =Nothing
End Sub

Public Sub CA001_OnClick()

    On Error GoTo erro:

	  Dim psOrigem As String
	  Dim piHandleOrigem As Long
	  Dim psXMLAtualizacao As String
	  Dim psXMLExclusao As String
	  Dim psMensagem As String
	  Dim piRetorno As Integer

	  '7/30/2010 11:17:06 AM - piRetorno = dllBSBen021.Beneficiario(CurrentSystem, 30414, False, psXMLAtualizacao, psXMLExclusao, 0, 0, 0, ))

	  psXMLAtualizacao = CStr( "<enderecos> <RegEndereco> <HANDLE>2083</HANDLE> <CEP>87070-260</CEP> <ESTADO>16</ESTADO> <MUNICIPIO>6056</MUNICIPIO> <BAIRRO>JD MONTREAL</BAIRRO> <TIPOLOGRADOURO>566</TIPOLOGRADOURO> <LOGRADOURO>COLOMBO 8678</LOGRADOURO> <NUMERO>213</NUMERO> <COMPLEMENTO/> <TELEFONE1/> <TELEFONE2/> <FAX/> <CELULAR/> <RAMAL/> <TIPOCOMPLEMENTO/> <ENDERECO1>S</ENDERECO1> <ENDERECO2>N</ENDERECO2> <ENDERECO3>S</ENDERECO3> </RegEndereco> </enderecos>" )
	  psXMLExclusao = ""

      Dim dllBSBen021    As Object
      Dim viHEndereco1 As Long
      Dim viHEndereco2 As Long
      Dim viHEndereco3 As Long

      Set dllBSBen021 = CreateBennerObject("BSBen021.AtualizacaoEndereco")

      If psXMLAtualizacao <> "" Then
      	psXMLAtualizacao = Replace( Replace( psXMLAtualizacao, "&lt", ">" ), "&gt", "<")
	  End If

	  If psXMLExclusao <> "" Then
	  	psXMLExclusao = Replace( Replace( psXMLExclusao, "&lt", ">" ), "&gt", "<")
	  End If

        piRetorno = dllBSBen021.Beneficiario(CurrentSystem, _
                                                                    30414, _
                                                                    False, _
                                                                    psXMLAtualizacao, _
                                                                    psXMLExclusao, _
                                                                    viHEndereco1, _
                                                                    viHEndereco2, _
                                                                    viHEndereco3, _
                                                                    psMensagem)


	 gravaAbrev("vihendereco1 = " + CStr( viHEndereco1) + _
				Chr(13) + "vihendereco2 = " + CStr( viHEndereco2) + _
				Chr(13) + "vihendereco3 = " + CStr( viHEndereco3))

	  gravaAbrev("psmensagem = "+ CStr( psMensagem ))
	  gravaAbrev("piRetorno = " + CStr(CLng( piRetorno )))


	  Set dllBSBen021 = Nothing

      Exit Sub

	erro:
		psMensagem = Err.Description
		piRetorno = 1
		ServiceVar("psMensagem") = CStr( psMensagem )
        ServiceVar("piRetorno") = CLng( piRetorno )

        Set dllBSBen021 = Nothing








Exit Sub

Dim qdll As Object
  Set qdll = CreateBennerObject("descobre.rotinas")
  qdll.rotina(CurrentSystem,1)
  Set qdll = Nothing

End Sub

Public Function SubstEstrutura(EstOriginal As String, CharSubst As String) As String
'  Dim EstNova As String

 ' EstOriginal = InputBox("Digite a estrutura original !", "Estrutura")

'  CharSubst = InputBox("Digite o novo caracter !", "novo caracter")

'EstNova = Mid(EstOriginal, 2, Len(EstOriginal))

  SubstEstrutura = Mid(EstOriginal, Len(CharSubst) + 1, Len(EstOriginal))

  SubstEstrutura = CharSubst + SubstEstrutura

End Function



Public Sub CA002_OnClick()'Fazer o update nas tabelas SFN_ROTINAARQUIVO E SFN_ROTINADOC -SMS: 20369 Coelho 07/04/2004

Dim qSel As Object
Dim qInsert As Object
Set qInsert = NewQuery
Set qSel = NewQuery

qSel.Add("SELECT HANDLE")
qSel.Add("FROM Z_TABELAS")
qSel.Add("WHERE NOME LIKE '" + CurrentQuery.FieldByName("ABREVIATURA").AsString + "%'")
qSel.Active = True

qInsert.Add("INSERT INTO W_D2WTABELAS")
qInsert.Add("(HANDLE, TABELA, CRIAR, OPCAORESULTADO)")
qInsert.Add("VALUES")
qInsert.Add("(:HANDLE, :HTABELA, 'N', 1)")

StartTransaction

While Not qSel.EOF
  qInsert.ParamByName("HANDLE").Value = NewHandle("W_D2WTABELAS")
  qInsert.ParamByName("HTABELA").Value =qSel.FieldByName("HANDLE").AsInteger
  qInsert.ExecSQL

  qSel.Next
Wend

Commit

MsgBox ("Processo concluído")

Exit Sub

  Dim qSelRotArq As Object
  Dim qSelRotDoc As Object
  Dim qUpRotArq As Object
  Dim qUpRotDoc As Object
  '--------------
  Set qSelRotArq = NewQuery
  Set qSelRotDoc = NewQuery
  Set qUpRotArq = NewQuery
  Set qUpRotDoc = NewQuery

  StartTransaction

  qSelRotArq.Clear
  qSelRotArq.Add("SELECT HANDLE FROM SFN_ROTINAARQUIVO ORDER BY HANDLE")
  qSelRotArq.Active = True

  While Not qSelRotArq.EOF
    qUpRotArq.Clear
    qUpRotArq.Add("UPDATE SFN_ROTINAARQUIVO SET CODIGO=:HROTARQ                     ")
    qUpRotArq.Add(" WHERE HANDLE=:HROTARQ                                          ")
    qUpRotArq.ParamByName("HROTARQ").AsInteger = qSelRotArq.FieldByName("HANDLE").AsInteger
    qUpRotArq.ExecSQL
    qSelRotArq.Next
  Wend

  qSelRotDoc.Clear
  qSelRotDoc.Add("SELECT HANDLE FROM SFN_ROTINADOC")
  qSelRotDoc.Active = True

  While Not qSelRotDoc.EOF
    qUpRotDoc.Clear
    qUpRotDoc.Add("UPDATE SFN_ROTINADOC SET CODIGO=:HROTDOC                        ")
    qUpRotDoc.Add(" WHERE HANDLE=:HROTDOC                                          ")
    qUpRotDoc.ParamByName("HROTDOC").AsInteger = qSelRotDoc.FieldByName("HANDLE").AsInteger
    qUpRotDoc.ExecSQL
    qSelRotDoc.Next
  Wend

  Commit

  Set qSelRotArq = Nothing
  Set qSelRotDoc = Nothing
  Set qUpRotArq = Nothing
  Set qUpRotDoc = Nothing

End Sub

Public Sub CA009_OnClick()

  Dim SQL As Object
  Dim SQL1 As Object
  Dim SQL2 As Object
  Dim vEstruturaSemMascara As String
  Dim vIndice As Integer
  Dim vDigitoVerificador As Integer
  Dim vQtdeNumerosMascara As Integer

  Set SQL = NewQuery
  Set SQL1 = NewQuery
  Set SQL2 = NewQuery

  SQL.Add("SELECT * FROM Z_MASCARAs WHERE TABELA = (SELECT HANDLE FROM Z_TABELAS WHERE NOME = 'SAM_TGE')")
  SQL.Active = True

  SQL1.Add("UPDATE SAM_TGE SET ESTRUTURA =:ESTRUTURA, ESTRUTURANUMERICA =:ESTRUTURANUMERICA WHERE HANDLE =:HANDLE ")

  SQL2.Add("SELECT HANDLE, ESTRUTURA, ULTIMONIVEL FROM SAM_TGE WHERE ULTIMONIVEL = 'S' AND HANDLE = 356")
  SQL2.Active = True

  If Not InTransaction Then StartTransaction

  While Not SQL2.EOF

    vEstruturaSemMascara = ""
    vQtdeNumerosMascara = 0

    For vIndice = 1 To Len(SQL2.FieldByName("ESTRUTURA").AsString)
      If InStr("0123456789", Mid(SQL2.FieldByName("ESTRUTURA").AsString, vIndice, 1))>0 Then
        vEstruturaSemMascara = vEstruturaSemMascara + _
                               Mid(SQL2.FieldByName("ESTRUTURA").AsString, vIndice, 1)
      End If
    Next vIndice

    For vIndice = 1 To Len(SQL.FieldByName("MASCARA").AsString)
      If InStr("0123456789", Mid(SQL.FieldByName("MASCARA").AsString, vIndice, 1))>0 Then
        vQtdeNumerosMascara = vQtdeNumerosMascara + 1
      End If
    Next vIndice

    If Not SQL.EOF Then
      vDigitoVerificador = Val(Modulo11(Mid(vEstruturaSemMascara, 1, Len(vEstruturaSemMascara) -1)))
      vEstruturaSemMascara = Left((vEstruturaSemMascara + "0000000000"), vQtdeNumerosMascara -1)
      vEstruturaSemMascara = Trim(vEstruturaSemMascara) + Trim(Str(vDigitoVerificador))
    End If



    SQL1.ParamByName("ESTRUTURA").Value = Mid(SQL2.FieldByName("ESTRUTURA").AsString, 1, Len(SQL2.FieldByName("ESTRUTURA").AsString) -1) + Trim(Str(vDigitoVerificador))

    SQL1.ParamByName("ESTRUTURANUMERICA").Value = Val(vEstruturaSemMascara)
    SQL1.ParamByName("HANDLE").Value = SQL2.FieldByName("HANDLE").AsInteger
    SQL1.ExecSQL

    SQL2.Next

  Wend

  Commit

  Set SQL = Nothing
  Set SQL1 = Nothing
  Set SQL2 = Nothing

End Sub


Declare Function EncryptPswd Lib "CMVISM32" (ByVal pswd$, ByVal encrypt$) As Integer


Public Sub CA010_OnClick()

Dim ret As String
Dim pwlen As Integer
Dim CacheFactory As Object


Set CacheFactory = CreateObject("CacheObject.Factory")

' preallocate the buffer, 8 nulls
ret$ = String$(8,Chr(0))

' encrypt the password for the connection string
pwlen = EncryptPswd("SYS",ret$)

' reset the length of the password string
' to match that of the encrypted password
ret$ = Left$(ret$,pwlen)

' Establish connection to server
Dim connectstring As String
connectstring = "cn_iptcp:192.168.1.218[1972]:TJDF:" & "_SYSTEM" & ":" & ret$


MsgBox(ret$)

Dim success As Boolean
success = CacheFactory.Connect(connectstring)

If success <> True Then
    Dim MyMsg As String
    MyMsg = "Connection failed."
    MsgBox MyMsg
End If

MsgBox(ret$)

End Sub



Public Sub CA013_OnClick()
  'Dim obj As Object
  'Set obj = CreateBennerObject("samutil.Atualizacao")
  'obj.AtualizaRegrapag(CurrentSystem)
  'Set obj = Nothing
  'Dim vDLLExportacaoOrizon As Object
  'Set vDLLExportacaoOrizon = CreateBennerObject("Benner.Saude.Orizon.Exportacao.ExportacaoOrizon")
  'vDLLExportacaoOrizon.Teste(CurrentSystem)
  'Set vDLLExportacaoOrizon = Nothing

  Dim SQL As Object
  Dim SQLUpd As Object

  Set SQL = NewQuery
  Set SQLUpd = NewQuery

  SQL.Add("SELECT HANDLE, ATRIBUTOS FROM W_PAGINAWIDGETS")
  SQLUpd.Add("UPDATE W_PAGINAWIDGETS SET ATRIBUTOS = :ATRIBUTOS WHERE HANDLE = :HANDLE")

  SQL.Active = True

  While Not SQL.EOF
    SQLUpd.ParamByName("HANDLE").AsInteger = SQL.FieldByName("HANDLE").AsInteger
    SQLUpd.ParamByName("ATRIBUTOS").AsString =  Replace(SQL.FieldByName("ATRIBUTOS").AsString, ">desaglocal<", ">QUALIDADEAG311<")
    SQLUpd.ExecSQL
    SQL.Next
  Wend

  Set SQL = Nothing
  Set SQLUpd = Nothing

  MsgBox("Terminou.")

End Sub


Public Sub CA014_OnClick()

  Dim Interface As Object
  Set Interface = CreateBennerObject("FINANCEIRO.Fatura")

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT HANDLE FROM SFN_FATURA")
  While Not SQL.EOF
    Interface.Atualiza(CurrentSystem, SQL.FieldByName("HANDLE").AsInteger)
    SQL.Next
  Wend

  '  Interface.Atualiza(CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set Interface = Nothing







  Set x = CreateBennerObject("CA014.Autorizacao")
  x.Exec(CurrentSystem, 0, 0, 0)'handle atendimento,autorizacao,documento
  'x.Consulta(366)
  Set x = Nothing
End Sub

Public Sub CA015_OnClick()




  Dim SQL As Object
  Dim SqlUp As Object

  Set SQL = NewQuery
  Set SqlUp = NewQuery

  StartTransaction

  SQL.Clear
  SQL.Add("SELECT C.HANDLE, A.CCNOME, A.CCCPFCNPJ                                     ")
  SQL.Add("  FROM SFN_CONTAFIN_ALTERACAO A,                                           ")
  SQL.Add("       SFN_CONTAFIN C                                                      ")
  SQL.Add(" WHERE C.HANDLE = A.CONTAFINANCEIRA                                        ")
  SQL.Add("   AND A.HANDLE IN (SELECT MAX(T.HANDLE)                                   ")
  SQL.Add("                      FROM SFN_CONTAFIN_ALTERACAO T                        ")
  SQL.Add("                     WHERE T.CONTAFINANCEIRA = A.CONTAFINANCEIRA           ")
  SQL.Add("                       AND T.SITUACAO = 'C'                                ") '-- Apenas as Confirmadas, desconsidera as solicitadas pois ainda podem ser processadas
  SQL.Add("                       AND ((T.CCNOME IS NULL) OR (T.CCCPFCNPJ IS NULL)))  ")
 ' Sql.Add("AND C.HANDLE = 37933                                                       ")

  SQL.Active = True


  While Not SQL.EOF
    SqlUp.Clear
    SqlUp.Add("UPDATE SFN_CONTAFIN     ")
    SqlUp.Add("   SET CCNOME = :PNOME, ")
    SqlUp.Add("       CCCPFCNPJ =:PCPF ")
    SqlUp.Add(" WHERE HANDLE =:PHANDLE ")
    SqlUp.ParamByName("PNOME").AsString = SQL.FieldByName("CCNOME").AsString
    SqlUp.ParamByName("PCPF").AsString = SQL.FieldByName("CCCPFCNPJ").AsString
    SqlUp.ParamByName("PHANDLE").AsInteger = SQL.FieldByName("HANDLE").AsInteger
    SqlUp.ExecSQL

    Sql.Next
  Wend

  Commit

  Set Sql = Nothing
  Set SqlUp = Nothing


'  Dim obj As Object
'  Dim vAux As String


'  Set obj = CreateBennerObject("SamSolicitAux.ExportaSolicitAux")
'  obj.Inicializar
'  Set obj = Nothing




  '  Dim vRotinaFinFat As Integer

  '  vAux =InputBox("Devolução de primeira mensalidade","Iforme o handle da SFN_ROTINAFINFAT")

  '  If vAux ="" Then
  '    MsgBox("Falta informar o handle")
  '    Exit Sub
  '  End If

  '  vRotinaFinFat =Val(vAux)

  '  Set Obj =CreateBennerObject("SamFaturamento.Faturamento")
  '  Obj.DevolucaoMensal(CurrentSystem,vRotinaFinFat)
  '  Set Obj =Nothing

  '  Dim Obj As Object

  '  Set Obj=CreateBennerObject("SamImpressao.Boleto")
  '  Obj.Inicializar
  '  Obj.ImprimirBoleto(10328)
  '  Obj.Finalizar
  '  Set Obj=Nothing
End Sub

Public Sub CA016_OnClick()
  'Set x =CreateBennerObject("Traduz.Rotinas")
  'x.Exec
  'Set x =Nothing
  Dim sqlBuscaDoc As Object
  Dim sqlLanc As Object
  Dim sqlDebCre As Object
  Dim sqlContLanc As Object
  Dim sqlFatLancCC As Object
  Dim sqlDelLanc As Object
  Dim upFatura As Object
  Dim upDocumento As Object
  '-----------------
  Set sqlBuscaDoc = NewQuery
  Set sqlLanc = NewQuery
  Set sqlDebCre = NewQuery
  Set sqlContLanc = NewQuery
  Set sqlFatLancCC = NewQuery
  Set sqlDelLanc = NewQuery
  Set upFatura = NewQuery
  Set upDocumento = NewQuery

  sqlBuscaDoc.Clear
  sqlBuscaDoc.Add("SELECT D.HANDLE DHANDLE, DF.FATURA FATURA				")
  sqlBuscaDoc.Add(" FROM SFN_DOCUMENTO D, SFN_DOCUMENTO_FATURA DF			")
  sqlBuscaDoc.Add("WHERE DF.DOCUMENTO = D.HANDLE							")
  sqlBuscaDoc.Add("AND D.NUMERO IN (2080886)						")

  sqlBuscaDoc.Active = True

  StartTransaction

  While Not sqlBuscaDoc.EOF

    sqlLanc.Clear
    sqlLanc.Add("SELECT FL.HANDLE	HANDLE						")
    sqlLanc.Add("FROM SFN_FATURA_LANC FL						")
    sqlLanc.Add("WHERE FL.HANDLE NOT IN (SELECT FL.HANDLE		")
    sqlLanc.Add("FROM SFN_FATURA_LANC FL						")
    sqlLanc.Add("WHERE FL.FATURA =:FATURA						")
    sqlLanc.Add("AND FL.TIPOLANCFIN = 14						")
    sqlLanc.Add("AND FL.OPERACAO = 6)							")
    sqlLanc.Add("AND FL.FATURA =:FATURA 						")
    sqlLanc.ParamByName("FATURA").AsInteger = sqlBuscaDoc.FieldByName("FATURA").AsInteger
    sqlLanc.Active = True

    While Not sqlLanc.EOF
      'Deletar os contabLancDEBCRE
      sqlDebCre.Clear
      sqlDebCre.Add("DELETE FROM SFN_CONTAB_LANC_DEBCRE                                  ")
      sqlDebCre.Add("WHERE HANDLE IN (SELECT E.HANDLE                                    ")
      sqlDebCre.Add("                   FROM SFN_CONTAB_LANC_DEBCRE E, SFN_CONTAB_LANC D ")
      sqlDebCre.Add("                  WHERE E.CONTABLANC = D.HANDLE                     ")
      sqlDebCre.Add("                    AND D.FATURALANC =:HFATURALANC)                 ")
      sqlDebCre.ParamByName("HFATURALANC").AsInteger = sqlLanc.FieldByName("HANDLE").AsInteger

      'Deletar os contabLanc
      sqlContLanc.Clear
      sqlContLanc.Add("DELETE FROM SFN_CONTAB_LANC           ")
      sqlContLanc.Add("WHERE HANDLE IN(SELECT D.HANDLE       ")
      sqlContLanc.Add("               FROM SFN_CONTAB_LANC D ")
      sqlContLanc.Add("WHERE D.FATURALANC =:HFATURALANC)     ")
      sqlContLanc.ParamByName("HFATURALANC").AsInteger = sqlLanc.FieldByName("HANDLE").AsInteger

      'Deletar os faturaLancCC
      sqlFatLancCC.Clear
      sqlFatLancCC.Add("DELETE FROM SFN_FATURA_LANC_CC                     ")
      sqlFatLancCC.Add(" WHERE HANDLE IN(SELECT D.HANDLE                   ")
      sqlFatLancCC.Add("                  FROM SFN_FATURA_LANC_CC D        ")
      sqlFatLancCC.Add("                 WHERE D.LANCAMENTO =:HFATURALANC) ")
      sqlFatLancCC.ParamByName("HFATURALANC").AsInteger = sqlLanc.FieldByName("HANDLE").AsInteger

      'Deletar os lancamentos da fatura
      sqlDelLanc.Clear
      sqlDelLanc.Add("DELETE                  ")
      sqlDelLanc.Add("FROM SFN_FATURA_LANC    ")
      sqlDelLanc.Add("WHERE HANDLE =:HLANC    ")
      sqlDelLanc.ParamByName("HLANC").AsInteger = sqlLanc.FieldByName("HANDLE").AsInteger

      sqlDebCre.ExecSQL
      sqlContLanc.ExecSQL
      sqlFatLancCC.ExecSQL
      sqlDelLanc.ExecSQL

      sqlLanc.Next
    Wend

    'Fazendo upDate na fatura para o Status aberta
    upFatura.Clear
    upFatura.Add("UPDATE SFN_FATURA SET SALDO=VALOR, BAIXADATA=NULL, BAIXAVALOR=NULL, SITUACAO ='A',            ")
    upFatura.Add("                      BAIXAJURO=NULL, BAIXAMULTA=NULL, BAIXACORRECAO=NULL, BAIXADESCONTO=NULL ")
    upFatura.Add(" WHERE HANDLE=:FATURA                                                                         ")
    upFatura.ParamByName("FATURA").AsInteger = sqlBuscaDoc.FieldByName("FATURA").AsInteger
    upFatura.ExecSQL

    upDocumento.Clear
    upDocumento.Add("UPDATE SFN_DOCUMENTO SET BAIXAVALOR=NULL, BAIXADESCONTO=NULL, BAIXAJURO=NULL, BAIXAMULTA=NULL, BAIXADATA=NULL")
    upDocumento.Add("WHERE HANDLE=:HANDLE")
    upDocumento.ParamByName("HANDLE").AsInteger = sqlBuscaDoc.FieldByName("DHANDLE").AsInteger
    upDocumento.ExecSQL

    sqlBuscaDoc.Next

  Wend

  Commit

  Set sqlBuscaDoc = Nothing
  Set sqlLanc = Nothing
  Set sqlDebCre = Nothing
  Set sqlContLanc = Nothing
  Set sqlFatLancCC = Nothing
  Set sqlDelLanc = Nothing
  Set upFatura = Nothing
  Set upDocumento = Nothing

End Sub

Public Sub CA017_OnClick()'Coelho -apagar documentos gerados por rotina documento,na CASSI.19/01/2004
  Dim SQL As Object
  Dim SQL1 As Object
  Dim SQL2 As Object
  Dim SQL3 As Object

  Set SQL = NewQuery
  Set SQL1 = NewQuery
  Set SQL2 = NewQuery
  Set SQL3 = NewQuery


  SQL.Clear
  'Query que busca os documentos a serem cancelados
  SQL.Add(" SELECT A.HANDLE ")
  SQL.Add(" FROM SFN_DOCUMENTO A, ")
  SQL.Add("      SFN_DOCUMENTO_FATURA B, ")
  SQL.Add("      SFN_FATURA C ")
  SQL.Add(" WHERE A.HANDLE = B.DOCUMENTO ")
  SQL.Add(" AND C.HANDLE = B.FATURA ")
  SQL.Add(" AND C.COMPETENCIA BETWEEN '01/01/2004' AND '31/01/2004' ")
  SQL.Add(" AND C.SITUACAO = 'A' ")
  SQL.Add(" AND A.DATAVENCIMENTO <= '01/01/2004' ")
  SQL.Add(" AND NOT EXISTS (SELECT D.DOCUMENTO  ")
  SQL.Add("                   FROM SFN_ROTINAARQUIVO_DOC D ")
  SQL.Add(" 		  WHERE A.HANDLE = D.DOCUMENTO) ")
  SQL.Active = True

  StartTransaction

  While Not SQL.EOF
    SQL1.Clear
    SQL1.Add("UPDATE SFN_DOCUMENTO SET ROTINADOC = NULL WHERE HANDLE =:HDOC         ")
    SQL1.ParamByName("HDOC").AsInteger = SQL.FieldByName("HANDLE").AsInteger
    SQL1.ExecSQL

    SQL2.Clear
    SQL2.Add("DELETE FROM SFN_DOCUMENTO_FATURA WHERE DOCUMENTO =:HDOC  ")
    SQL2.ParamByName("HDOC").AsInteger = SQL.FieldByName("HANDLE").AsInteger
    SQL2.ExecSQL

    SQL3.Clear
    SQL3.Add("UPDATE SFN_DOCUMENTO                                                  ")
    SQL3.Add("Set CANCDATA = '2004-01-19 00:00:00',                                 ")
    SQL3.Add("    CANCMOTIVO = 'Gerado com data contabil incorreta' ,               ")
    SQL3.Add("    FATURADESPESA = Null                                              ")
    SQL3.Add("WHERE HANDLE =:HDOC")
    SQL3.ParamByName("HDOC").AsInteger = SQL.FieldByName("HANDLE").AsInteger
    SQL3.ExecSQL
    SQL.Next
  Wend

  Commit

  Set SQL = Nothing
  Set SQL1 = Nothing
  Set SQL2 = Nothing
  Set SQL3 = Nothing

End Sub

Public Sub CA018_OnClick()
  Dim QContas As Object
  Set QContas = NewQuery


  Dim VERIFICA As Object
  Dim CONTAFIN As Object
  Dim ALTERA As Object
  Dim CODBANCO As Object
  Dim CODAGENCIA As Object
  Dim CODOPERADORA As Object
  Set VERIFICA = NewQuery
  Set CONTAFIN = NewQuery
  Set ALTERA = NewQuery
  Set CODBANCO = NewQuery
  Set CODAGENCIA = NewQuery
  Set CODOPERADORA = NewQuery
  Dim DADOSANTERIORES As String
  Dim tipodoc As Object
  Set tipodoc = NewQuery
  Dim SQLCLASSECONTABIL As Object
  Set SQLCLASSECONTABIL = NewQuery
  Dim vClasseContabil As String



  QContas.Active = False
  QContas.Clear
  QContas.Add("Select * ")
  QContas.Add("  FROM SFN_CONTAFIN_ALTERACAO ")
  QContas.Add(" WHERE SITUACAO = 'S' ")
  QContas.Add("   And TABRESPONSAVEL = 2 ")
  QContas.Add(" ORDER BY HANDLE ")
  QContas.Active = True
  QContas.First
  While Not (QContas.EOF)


    'Seleciona o registro da Tabela SFN_CONTAFIN
    VERIFICA.Clear
    VERIFICA.Add("SELECT * FROM SFN_CONTAFIN")
    VERIFICA.Add("WHERE HANDLE=:HANDLE")
    VERIFICA.ParamByName("HANDLE").AsInteger = QContas.FieldByName("CONTAFINANCEIRA").AsInteger
    VERIFICA.Active = True

    If Not VERIFICA.FieldByName("CLASSECONTABIL").IsNull Then
      SQLCLASSECONTABIL.Clear
      SQLCLASSECONTABIL.Active = False
      SQLCLASSECONTABIL.Add("SELECT ESTRUTURA, DESCRICAO FROM SFN_CLASSECONTABIL WHERE HANDLE=:HANDLE")
      SQLCLASSECONTABIL.ParamByName("HANDLE").AsInteger = VERIFICA.FieldByName("CLASSECONTABIL").AsInteger
      SQLCLASSECONTABIL.Active = True

      vClasseContabil = SQLCLASSECONTABIL.FieldByName("ESTRUTURA").AsString + " - " + SQLCLASSECONTABIL.FieldByName("DESCRICAO").AsString
    Else
      vClasseContabil = ""
    End If

    DADOSANTERIORES = ""

    'Testa se a alteraçao é do tipo Conta corrente
    If VERIFICA.FieldByName("TABGERACAOPAG").AsInteger = 1 Then
      DADOSANTERIORES = "Geração pagamento  = Conta corrente"
    End If

    If VERIFICA.FieldByName("TABGERACAOREC").AsInteger = 1 Then
      If DADOSANTERIORES <> "" Then
        DADOSANTERIORES = DADOSANTERIORES + Chr(13) + _
                          "Geração recebimento = Conta corrente"
      Else
        DADOSANTERIORES = "Geração recebimento = Conta corrente"
      End If
    End If

    If VERIFICA.FieldByName("TABGERACAOPAG").AsInteger = 1 Or VERIFICA.FieldByName("TABGERACAOREC").AsInteger = 1Then

      'Busca o número do Banco
      CODBANCO.Clear
      CODBANCO.Add("SELECT CODIGO FROM SFN_BANCO WHERE HANDLE = :PBANCO")
      CODBANCO.ParamByName("PBANCO").AsInteger = VERIFICA.FieldByName("BANCO").AsInteger
      CODBANCO.Active = True

      'Busca o número da Agência
      CODAGENCIA.Clear
      CODAGENCIA.Add("SELECT AGENCIA FROM SFN_AGENCIA WHERE HANDLE = :PAGENCIA")
      CODAGENCIA.ParamByName("PAGENCIA").Value = VERIFICA.FieldByName("AGENCIA").AsInteger
      CODAGENCIA.Active = True

      'Registra os dados anteriores na tabela de alteraçoes
      DADOSANTERIORES = DADOSANTERIORES + Chr(13) + _
                        "Banco = " + CODBANCO.FieldByName("CODIGO").AsString + " | " + _
                        "Agência = " + CODAGENCIA.FieldByName("AGENCIA").AsString + " | " + _
                        "Conta Corrente = " + VERIFICA.FieldByName("CONTACORRENTE").AsString + " | " + _
                        "DV = " + VERIFICA.FieldByName("DV").AsString + " | " + _
                        "Nome = " + VERIFICA.FieldByName("CCNOME").AsString + " | " + _
                        "CPF = " + VERIFICA.FieldByName("CCCPFCNPJ").AsString + " | " + _
                        "Não Gerar Documento = " + VERIFICA.FieldByName("NAOGERARDOCUMENTO").AsString + " | " + _
                        "Não Cobrar Tarifa = " + VERIFICA.FieldByName("NAOCOBRARTARIFA").AsString


      'ALTERA.Clear
      'ALTERA.Add("UPDATE SFN_CONTAFIN_ALTERACAO Set DADOSANTERIORES = :PDADOS WHERE HANDLE = :PHANDLE")
      'ALTERA.ParamByName("PDADOS").Value = DADOSANTERIORES
      'ALTERA.ParamByName("PHANDLE").Value = QContas.FieldByName("HANDLE").AsInteger
      'ALTERA.ExecSQL

    End If


    'Testa se a alteração é do tipo cartao de credito
    If VERIFICA.FieldByName("TABGERACAOPAG").AsInteger = 2 Then
      If DADOSANTERIORES <> "" Then
        DADOSANTERIORES = DADOSANTERIORES + Chr(13) + _
                          "Geração pagamento  = Cartão de crédito"
      Else
        DADOSANTERIORES = "Geração pagamento  = Cartão de crédito"
      End If
    End If

    If VERIFICA.FieldByName("TABGERACAOREC").AsInteger = 2 Then
      If DADOSANTERIORES <> "" Then
        DADOSANTERIORES = DADOSANTERIORES + Chr(13) + _
                          "Geração recebimento  = Cartão de crédito"
      Else
        DADOSANTERIORES = "Geração recebimento  = Cartão de crédito"
      End If
    End If

    If VERIFICA.FieldByName("TABGERACAOPAG").AsInteger = 2 Or VERIFICA.FieldByName("TABGERACAOREC").AsInteger = 2 Then

      'vERIFICA O NOME DA OPERADORA DO CARTAO
      CODOPERADORA.Clear
      CODOPERADORA.Add("SELECT CODIGO, NOME")
      CODOPERADORA.Add("  FROM SFN_CARTAOOPERADORA")
      CODOPERADORA.Add(" WHERE HANDLE = :HANDLE")
      CODOPERADORA.ParamByName("HANDLE").Value = VERIFICA.FieldByName("CARTAOOPERADORA").AsInteger
      CODOPERADORA.Active = True

      DADOSANTERIORES = DADOSANTERIORES + Chr(13) + _
                        "Código operadora = " + CODOPERADORA.FieldByName("CODIGO").AsString + " | " + _
                        "Nome operadora = " + CODOPERADORA.FieldByName("NOME").AsString + " | " + _
                        "Número do cartão = " + VERIFICA.FieldByName("CARTAOCREDITO").AsString + " | " + _
                        "Validade = " + VERIFICA.FieldByName("CARTAOVALIDADE").AsString + " | " + _
                        "Não gerar documento = " + VERIFICA.FieldByName("NAOGERARDOCUMENTO").AsString + " | " + _
                        "Não cobrar tarifa = " + VERIFICA.FieldByName("NAOCOBRARTARIFA").AsString

      'Registra os dados anteriores na tabela de alteraçoes
      'ALTERA.Clear
      'ALTERA.Add("UPDATE SFN_CONTAFIN_ALTERACAO Set DADOSANTERIORES = :PDADOS WHERE HANDLE = :PHANDLE")
      'ALTERA.ParamByName("PDADOS").Value = DADOSANTERIORES
      'ALTERA.ParamByName("PHANDLE").Value = QContas.FieldByName("HANDLE").AsInteger
      'ALTERA.ExecSQL
    End If

    'Testa se a alteração é do tipo Título
    If VERIFICA.FieldByName("TABGERACAOPAG").AsInteger = 3 Then

      tipodoc.Active = False
      tipodoc.Clear
      tipodoc.Add("SELECT DESCRICAO FROM SFN_TIPODOCUMENTO")
      tipodoc.Add(" WHERE HANDLE = :HANDLEDOC")
      tipodoc.ParamByName("HANDLEDOC").Value = VERIFICA.FieldByName("TIPODOCUMENTOPAG").AsInteger
      tipodoc.Active = True

      If DADOSANTERIORES <> "" Then
        DADOSANTERIORES = DADOSANTERIORES + Chr(13) + _
                          "Geração pagamento  = Título" + Chr(13) + _
                          "Tipo do documento = " + tipodoc.FieldByName("DESCRICAO").AsString + " | " + _
                          "Não Gerar Documento = " + VERIFICA.FieldByName("NAOGERARDOCUMENTO").AsString + " | " + _
                          "Não Cobrar Tarifa = " + VERIFICA.FieldByName("NAOCOBRARTARIFA").AsString
      Else
        DADOSANTERIORES = "Geração pagamento  = Título" + Chr(13) + _
                          "Tipo do documento = " + tipodoc.FieldByName("DESCRICAO").AsString + " | " + _
                          "Não Gerar Documento = " + VERIFICA.FieldByName("NAOGERARDOCUMENTO").AsString + " | " + _
                          "Não Cobrar Tarifa = " + VERIFICA.FieldByName("NAOCOBRARTARIFA").AsString
      End If
    End If

    If VERIFICA.FieldByName("TABGERACAOREC").AsInteger = 3 Then

      tipodoc.Active = False
      tipodoc.Clear
      tipodoc.Add("SELECT DESCRICAO FROM SFN_TIPODOCUMENTO")
      tipodoc.Add(" WHERE HANDLE = :HANDLEDOC")
      tipodoc.ParamByName("HANDLEDOC").Value = VERIFICA.FieldByName("TIPODOCUMENTOREC").AsInteger
      tipodoc.Active = True

      If DADOSANTERIORES <> "" Then
        DADOSANTERIORES = DADOSANTERIORES + Chr(13) + _
                          "Geração recebimento = Título" + Chr(13) + _
                          "Tipo do documento = " + tipodoc.FieldByName("DESCRICAO").AsString + " | " + _
                          "Não Gerar Documento = " + VERIFICA.FieldByName("NAOGERARDOCUMENTO").AsString + " | " + _
                          "Não Cobrar Tarifa = " + VERIFICA.FieldByName("NAOCOBRARTARIFA").AsString
      Else
        DADOSANTERIORES = "Geração recebimento = Título" + Chr(13) + _
                          "Tipo do documento = " + tipodoc.FieldByName("DESCRICAO").AsString + " | " + _
                          "Não Gerar Documento = " + VERIFICA.FieldByName("NAOGERARDOCUMENTO").AsString + " | " + _
                          "Não Cobrar Tarifa = " + VERIFICA.FieldByName("NAOCOBRARTARIFA").AsString
      End If
    End If

    On Error GoTo ERRO

    'StartTransaction

    If vClasseContabil <> "" Then
      DADOSANTERIORES = DADOSANTERIORES + Chr(13) + "Classe Contábil = " + vClasseContabil
    Else
      'DADOSANTERIORES = DADOSANTERIORES + Chr(13) + "Classe Contábil = " + vC
    End If

    'Registra os dados anteriores na tabela de alteraçoes
    ALTERA.Clear
    ALTERA.Add("UPDATE SFN_CONTAFIN_ALTERACAO Set DADOSANTERIORES = :PDADOS WHERE HANDLE = :PHANDLE")
    ALTERA.ParamByName("PDADOS").AsMemo = DADOSANTERIORES
    ALTERA.ParamByName("PHANDLE").Value = QContas.FieldByName("HANDLE").AsInteger
    ALTERA.ExecSQL

    'SMS 25916 - Kristian
    If QContas.FieldByName("TABRESPONSAVEL").AsInteger = 1 Then
      CONTAFIN.Clear
      CONTAFIN.Add("UPDATE SFN_CONTAFIN ")
      CONTAFIN.Add("   SET TABCOBRANCAOUTRORESPONSAVEL = :TABRESP,")
      CONTAFIN.Add("       BENEFICIARIODESTINOCOBRANCA = :BENEF   ")
      CONTAFIN.Add(" WHERE HANDLE = :PHANDLE")

      CONTAFIN.ParamByName("PHANDLE").AsInteger = QContas.FieldByName("CONTAFINANCEIRA").AsInteger
      CONTAFIN.ParamByName("TABRESP").AsInteger = QContas.FieldByName("TABCOBRANCAOUTRORESPONSAVEL").AsInteger
      CONTAFIN.ParamByName("BENEF").AsInteger = QContas.FieldByName("BENEFICIARIODESTINOCOBRANCA").AsInteger
      CONTAFIN.ExecSQL

    End If

    'FIM SMS 25916


    'sms 33910 atualiza dados da conta corrente independente do que estiver configurado nos campos TABGERACAOREC e TABGERACAOPAG
    If (Not QContas.FieldByName("BANCO").IsNull) Or (Not QContas.FieldByName("AGENCIA").IsNull) Or _
         (Not QContas.FieldByName("CONTACORRENTE").IsNull) Or (Not QContas.FieldByName("DV").IsNull) Or _
         (Not QContas.FieldByName("CCNOME").IsNull) Or (Not QContas.FieldByName("CCCPFCNPJ").IsNull) Then

      CONTAFIN.Clear
      CONTAFIN.Add("UPDATE SFN_CONTAFIN SET               				")
      If Not QContas.FieldByName("BANCO").IsNull Then
        CONTAFIN.Add("                        BANCO = :PBANCO,		    ")
        CONTAFIN.ParamByName("PBANCO").Value = QContas.FieldByName("BANCO").AsInteger
      End If
      If Not QContas.FieldByName("AGENCIA").IsNull Then
        CONTAFIN.Add("                        AGENCIA = :PAGENCIA,		")
        CONTAFIN.ParamByName("PAGENCIA").Value = QContas.FieldByName("AGENCIA").AsInteger
      End If
      If Not QContas.FieldByName("CONTACORRENTE").IsNull Then
        CONTAFIN.Add("                        CONTACORRENTE = :PCONTA,	")
        CONTAFIN.ParamByName("PCONTA").Value = QContas.FieldByName("CONTACORRENTE").AsString
      End If
      If Not QContas.FieldByName("DV").IsNull Then
        CONTAFIN.Add("                        DV = :PDV,                  ")
        CONTAFIN.ParamByName("PDV").Value = QContas.FieldByName("DV").AsString
      End If
      If Not QContas.FieldByName("CCNOME").IsNull Then
        CONTAFIN.Add("                        CCNOME = :PCCNOME,          ")
        CONTAFIN.ParamByName("PCCNOME").Value = QContas.FieldByName("CCNOME").AsString
      End If
      If Not QContas.FieldByName("CCCPFCNPJ").IsNull Then
        CONTAFIN.Add("                        CCCPFCNPJ = :PCCCPFCNPJ,     ")
        CONTAFIN.ParamByName("PCCCPFCNPJ").Value = QContas.FieldByName("CCCPFCNPJ").AsString
      End If
      CONTAFIN.Add("                          HANDLE = HANDLE     ")

      CONTAFIN.Add("WHERE HANDLE = :PHANDLE")
      CONTAFIN.ParamByName("PHANDLE").Value = QContas.FieldByName("CONTAFINANCEIRA").AsInteger
      CONTAFIN.ExecSQL
    End If 'sms 33910

    If QContas.FieldByName("TABGERACAOPAG").AsInteger = 1 Then

      'Atualiza as alterações na tabela conta financeira
      CONTAFIN.Clear
      CONTAFIN.Add("UPDATE SFN_CONTAFIN SET TABGERACAOPAG = 1,				")
      'CONTAFIN.Add("                        BANCO = :PBANCO,				") sms 33910
      'CONTAFIN.Add("                        AGENCIA = :PAGENCIA,			")
      'CONTAFIN.Add("                        CONTACORRENTE = :PCONTA,		")
      'CONTAFIN.Add("                        DV = :PDV,                      ")
      'CONTAFIN.Add("                        CCNOME = :PCCNOME,              ")
      'CONTAFIN.Add("                        CCCPFCNPJ = :PCCCPFCNPJ,        ")
      CONTAFIN.Add("                        NAOGERARDOCUMENTO = :PDOC,      ")
      If Not QContas.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.Add("                        NAOCOBRARTARIFA = :PTARIFA,   ")
        CONTAFIN.Add("                        CLASSECONTABIL = :CLASSECONTABIL")
      Else
        CONTAFIN.Add("                        NAOCOBRARTARIFA = :PTARIFA")
      End If
      CONTAFIN.Add("WHERE HANDLE = :PHANDLE")

      'CONTAFIN.ParamByName("PBANCO").Value    = QContas.FieldByName("BANCO").AsInteger sms 33910
      'CONTAFIN.ParamByName("PAGENCIA").Value  = QContas.FieldByName("AGENCIA").AsInteger
      'CONTAFIN.ParamByName("PCONTA").Value    = QContas.FieldByName("CONTACORRENTE").AsString
      'CONTAFIN.ParamByName("PDV").Value       = QContas.FieldByName("DV").AsString
      'CONTAFIN.ParamByName("PCCNOME").Value   = QContas.FieldByName("CCNOME").AsString
      'CONTAFIN.ParamByName("PCCCPFCNPJ").Value=QContas.FieldByName("CCCPFCNPJ").AsString
      CONTAFIN.ParamByName("PDOC").Value = QContas.FieldByName("NAOGERARDOCUMENTO").AsString
      CONTAFIN.ParamByName("PTARIFA").Value = QContas.FieldByName("NAOCOBRARTARIFA").AsString
      If Not QContas.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.ParamByName("CLASSECONTABIL").Value = QContas.FieldByName("CLASSECONTABIL").Value
      End If
      CONTAFIN.ParamByName("PHANDLE").Value = QContas.FieldByName("CONTAFINANCEIRA").AsInteger
      CONTAFIN.ExecSQL

    End If

    If QContas.FieldByName("TABGERACAOREC").AsInteger = 1 Then

      'Atualiza as alterações na tabela conta financeira
      CONTAFIN.Clear
      CONTAFIN.Add("UPDATE SFN_CONTAFIN SET TABGERACAOREC = 1,              ")
      'CONTAFIN.Add("                        BANCO = :PBANCO,                ") sms 33910
      'CONTAFIN.Add("                        AGENCIA = :PAGENCIA,            ")
      'CONTAFIN.Add("                        CONTACORRENTE = :PCONTA,        ")
      'CONTAFIN.Add("                        DV = :PDV,                      ")
      'CONTAFIN.Add("                        CCNOME = :PCCNOME,              ")
      'CONTAFIN.Add("                        CCCPFCNPJ = :PCCCPFCNPJ,        ")
      CONTAFIN.Add("                        NAOGERARDOCUMENTO = :PDOC,      ")
      If Not QContas.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.Add("                        NAOCOBRARTARIFA = :PTARIFA,     ")
        CONTAFIN.Add("                        CLASSECONTABIL = :CLASSECONTABIL")
      Else
        CONTAFIN.Add("                        NAOCOBRARTARIFA = :PTARIFA,")
        CONTAFIN.Add("                        CLASSECONTABIL = NULL")
      End If
      CONTAFIN.Add("WHERE HANDLE = :PHANDLE")
      'CONTAFIN.ParamByName("PBANCO").Value   = QContas.FieldByName("BANCO").AsInteger sms 33910
      'CONTAFIN.ParamByName("PAGENCIA").Value = QContas.FieldByName("AGENCIA").AsInteger
      'CONTAFIN.ParamByName("PCONTA").Value   = QContas.FieldByName("CONTACORRENTE").AsString
      'CONTAFIN.ParamByName("PDV").Value      = QContas.FieldByName("DV").AsString
      'CONTAFIN.ParamByName("PCCNOME").Value  = QContas.FieldByName("CCNOME").AsString
      'CONTAFIN.ParamByName("PCCCPFCNPJ").Value   = QContas.FieldByName("CCCPFCNPJ").AsString
      CONTAFIN.ParamByName("PDOC").Value = QContas.FieldByName("NAOGERARDOCUMENTO").AsString
      CONTAFIN.ParamByName("PTARIFA").Value = QContas.FieldByName("NAOCOBRARTARIFA").AsString
      CONTAFIN.ParamByName("PHANDLE").Value = QContas.FieldByName("CONTAFINANCEIRA").AsInteger
      If Not QContas.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.ParamByName("CLASSECONTABIL").Value = QContas.FieldByName("CLASSECONTABIL").Value
      End If
      CONTAFIN.ExecSQL
    End If


    If QContas.FieldByName("TABGERACAOPAG").AsInteger = 2 Then

      'Atualiza as alterações na tabela conta financeira
      CONTAFIN.Clear
      CONTAFIN.Add("UPDATE SFN_CONTAFIN SET TABGERACAOPAG = 2, CARTAOOPERADORA = :OPERADORA, CARTAOCREDITO = :PCARTAO,")
      CONTAFIN.Add("                        CARTAOVALIDADE = :PVALIDADE, NAOGERARDOCUMENTO = :PDOC,                   ")
      If Not QContas.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.Add("                      NAOCOBRARTARIFA = :PTARIFA, CLASSECONTABIL = :CLASSECONTABIL")
      Else
        CONTAFIN.Add("                      NAOCOBRARTARIFA = :PTARIFA,")
        CONTAFIN.Add("                        CLASSECONTABIL = NULL")
      End If
      CONTAFIN.Add("WHERE HANDLE = :PHANDLE")
      CONTAFIN.ParamByName("OPERADORA").Value = QContas.FieldByName("CARTAOOPERADORA").AsInteger
      CONTAFIN.ParamByName("PCARTAO").Value = QContas.FieldByName("CARTAOCREDITO").AsString
      CONTAFIN.ParamByName("PVALIDADE").Value = QContas.FieldByName("CARTAOVALIDADE").AsDateTime
      CONTAFIN.ParamByName("PDOC").Value = QContas.FieldByName("NAOGERARDOCUMENTO").AsString
      CONTAFIN.ParamByName("PTARIFA").Value = QContas.FieldByName("NAOCOBRARTARIFA").AsString
      CONTAFIN.ParamByName("PHANDLE").Value = QContas.FieldByName("CONTAFINANCEIRA").AsInteger
      If Not QContas.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.ParamByName("CLASSECONTABIL").Value = QContas.FieldByName("CLASSECONTABIL").Value
      End If
      CONTAFIN.ExecSQL
    End If

    If QContas.FieldByName("TABGERACAOREC").AsInteger = 2 Then

      'Atualiza as alterações na tabela conta financeira
      CONTAFIN.Clear
      CONTAFIN.Add("UPDATE SFN_CONTAFIN SET TABGERACAOREC = 2, CARTAOOPERADORA = :OPERADORA, CARTAOCREDITO = :PCARTAO,")
      CONTAFIN.Add("                        CARTAOVALIDADE = :PVALIDADE, NAOGERARDOCUMENTO = :PDOC, ")
      If Not QContas.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.Add("                      NAOCOBRARTARIFA = :PTARIFA, CLASSECONTABIL = :CLASSECONTABIL")
      Else
        CONTAFIN.Add("                      NAOCOBRARTARIFA = :PTARIFA,")
        CONTAFIN.Add("                        CLASSECONTABIL = NULL")
      End If
      CONTAFIN.Add("WHERE HANDLE = :PHANDLE")
      CONTAFIN.ParamByName("OPERADORA").Value = QContas.FieldByName("CARTAOOPERADORA").AsInteger
      CONTAFIN.ParamByName("PCARTAO").Value = QContas.FieldByName("CARTAOCREDITO").AsString
      CONTAFIN.ParamByName("PVALIDADE").Value = QContas.FieldByName("CARTAOVALIDADE").AsDateTime
      CONTAFIN.ParamByName("PDOC").Value = QContas.FieldByName("NAOGERARDOCUMENTO").AsString
      CONTAFIN.ParamByName("PTARIFA").Value = QContas.FieldByName("NAOCOBRARTARIFA").AsString
      CONTAFIN.ParamByName("PHANDLE").Value = QContas.FieldByName("CONTAFINANCEIRA").AsInteger
      If Not QContas.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.ParamByName("CLASSECONTABIL").Value = QContas.FieldByName("CLASSECONTABIL").Value
      End If
      CONTAFIN.ExecSQL
    End If


    If QContas.FieldByName("TABGERACAOPAG").AsInteger = 3 Then

      'Atualiza as alterações na tabela conta financeira
      CONTAFIN.Clear
      CONTAFIN.Add("UPDATE SFN_CONTAFIN SET TABGERACAOPAG = 3, TIPODOCUMENTOPAG = :TIPODOC,")
      CONTAFIN.Add("                        NAOGERARDOCUMENTO = :PDOC, ")
      If Not QContas.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.Add("                      NAOCOBRARTARIFA = :PTARIFA, CLASSECONTABIL = :CLASSECONTABIL")
      Else
        CONTAFIN.Add("                      NAOCOBRARTARIFA = :PTARIFA,")
        CONTAFIN.Add("                        CLASSECONTABIL = NULL")
      End If
      CONTAFIN.Add("WHERE HANDLE = :PHANDLE")
      CONTAFIN.ParamByName("TIPODOC").Value = QContas.FieldByName("TIPODOCUMENTOPAG").AsInteger
      CONTAFIN.ParamByName("PDOC").Value = QContas.FieldByName("NAOGERARDOCUMENTO").AsString
      CONTAFIN.ParamByName("PTARIFA").Value = QContas.FieldByName("NAOCOBRARTARIFA").AsString
      CONTAFIN.ParamByName("PHANDLE").Value = QContas.FieldByName("CONTAFINANCEIRA").AsInteger
      If Not QContas.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.ParamByName("CLASSECONTABIL").Value = QContas.FieldByName("CLASSECONTABIL").Value
      End If
      CONTAFIN.ExecSQL
    End If

    If QContas.FieldByName("TABGERACAOREC").AsInteger = 3 Then

      'Atualiza as alterações na tabela conta financeira
      CONTAFIN.Clear
      CONTAFIN.Add("UPDATE SFN_CONTAFIN SET TABGERACAOREC = 3, TIPODOCUMENTOREC = :TIPODOC,")
      CONTAFIN.Add("                        NAOGERARDOCUMENTO = :PDOC, ")
      If Not QContas.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.Add("                      NAOCOBRARTARIFA = :PTARIFA, CLASSECONTABIL = :CLASSECONTABIL")
      Else
        CONTAFIN.Add("                      NAOCOBRARTARIFA = :PTARIFA,")
        CONTAFIN.Add("                        CLASSECONTABIL = NULL")
      End If
      CONTAFIN.Add("WHERE HANDLE = :PHANDLE")
      CONTAFIN.ParamByName("TIPODOC").Value = QContas.FieldByName("TIPODOCUMENTOREC").AsInteger
      CONTAFIN.ParamByName("PDOC").Value = QContas.FieldByName("NAOGERARDOCUMENTO").AsString
      CONTAFIN.ParamByName("PTARIFA").Value = QContas.FieldByName("NAOCOBRARTARIFA").AsString
      CONTAFIN.ParamByName("PHANDLE").Value = QContas.FieldByName("CONTAFINANCEIRA").AsInteger
      If Not QContas.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.ParamByName("CLASSECONTABIL").Value = QContas.FieldByName("CLASSECONTABIL").Value
      End If
      CONTAFIN.ExecSQL
    End If


    'Grava a data e o usuário responsável pela confirmação
    ALTERA.Clear
    ALTERA.Add("UPDATE SFN_CONTAFIN_ALTERACAO SET SITUACAO = 'C', CONFIRMADODATA = :PDATA, CONFIRMADOUSUARIO = :PUSUARIO")
    ALTERA.Add("WHERE HANDLE = :PHANDLE")
    ALTERA.ParamByName("PDATA").Value = ServerDate
    ALTERA.ParamByName("PUSUARIO").Value = CurrentUser
    ALTERA.ParamByName("PHANDLE").Value = QContas.FieldByName("HANDLE").AsInteger
    ALTERA.ExecSQL

    'Commit


    'André - SMS 24062 - 27/08/2004

    Dim q As Object
    Dim q1 As Object
    Dim q2 As Object
    Dim q3 As Object
    Dim q4 As Object
    Dim q5 As Object
    Dim q6 As Object
    Dim q7 As Object

    Dim Resposta As Variant

    Set q = NewQuery
    Set q1 = NewQuery
    Set q2 = NewQuery
    Set q3 = NewQuery
    Set q4 = NewQuery
    Set q5 = NewQuery
    Set q6 = NewQuery
    Set q7 = NewQuery

    If QContas.FieldByName("TABGERACAOPAG").AsInteger = 1 Then
      If viBanco <> QContas.FieldByName("BANCO").Value Or _
                                         viAgencia <> QContas.FieldByName("AGENCIA").Value Or _
                                         VCC <> QContas.FieldByName("CONTACORRENTE").Value Or _
                                         vDV <> QContas.FieldByName("DV").Value Then

        q.Clear
        q.Add (" SELECT DISTINCT CONTAFINANCEIRA                                     ")
        q.Add ("   FROM SFN_FATURA FAT                                               ")
        q.Add ("        JOIN SFN_CONTAFIN CON On (FAT.CONTAFINANCEIRA = CON.HANDLE)  ")
        q.Add (" WHERE CON.PRESTADOR = :PRESTADOR                                    ")
        q.Text = SqlConverte(q.Text, SQLServer)
        MsgBox(q.Text)
        q.ParamByName("PRESTADOR").Value = RecordHandleOfTable("SAM_PRESTADOR")
        q.Active = True

        q1.Clear
        q1.Add (" SELECT DISTINCT DOC.HANDLE, DOC.CONTAFINANCEIRA, DOC.TIPODOCUMENTO, DOC.DATAEMISSAO,   ")
        q1.Add ("        DOC.DATAVENCIMENTO, DOC.COMPETENCIA, DOC.NUMERO, DOC.NOSSONUMERO,               ")
        q1.Add ("        DOC.TESOURARIA, DOC.ROTINADOC, DOC.REGRAFINANCEIRA, DOC.CONTRATO,               ")
        q1.Add ("        DOC.FAMILIA, DOC.IDENTIFICADORPAGAMENTO, DFAT.FATURA, DFAT.DOCUMENTO,           ")
        q1.Add ("        DFAT.SALDO, DFAT.NATUREZA, DFAT.VALORJURO, DFAT.VALORMULTA, DFAT.VALORCORRECAO, ")
        q1.Add ("        DFAT.VALORDESCONTO, DFAT.VALORTOTAL, FAT.SITUACAO, DOC.CANCDATA,                ")
        q1.Add ("        DOC.VALOR, DOC.NATUREZA NATU, DOC.VALORTOTAL VLTOT, DOC.VALORJURO JUROS,        ")
        q1.Add ("        DOC.VALORCORRECAO CORRECAO, DOC.VALORDESCONTO DESCONTO, DOC.VALORMULTA MULTA    ")
        q1.Add ("   FROM SFN_DOCUMENTO DOC,                                                              ")
        q1.Add ("        SFN_FATURA FAT,                                                                 ")
        q1.Add ("        SFN_DOCUMENTO_FATURA DFAT,                                                      ")
        q1.Add ("        SFN_ROTINAARQUIVO_DOC RDOC                                                      ")
        q1.Add ("  WHERE DOC.HANDLE = DFAT.DOCUMENTO                                                     ")
        q1.Add ("    AND FAT.HANDLE = DFAT.FATURA                                                        ")
        q1.Add ("    AND RDOC.DOCUMENTO = DOC.HANDLE                                                     ")
        q1.Add ("    AND RDOC.TABENVIORETORNO = 2                                                        ")
        q1.Add ("    AND DFAT.FATURA IN (SELECT HANDLE                                                   ")
        q1.Add ("                          FROM SFN_FATURA                                               ")
        q1.Add ("                         WHERE CONTAFINANCEIRA = :CONTAFIN)                             ")
        q1.Text = SqlConverte(q1.Text, SQLServer)
        q1.ParamByName("CONTAFIN").Value = q.FieldByName("CONTAFINANCEIRA").AsInteger
        q1.Active = True

        While Not (q1.EOF)
          If q1.FieldByName("SITUACAO").Value <> "A" Then
            Exit Sub
          End If
          q1.Next
        Wend
        q1.First
        If Not q1.EOF Then
          Resposta = MsgBox("Existem documentos em aberto. Deseja emitir novos documentos?", vbYesNo)
          If Resposta = vbYes Then
            While Not (q1.EOF)

              q7.Clear
              q7.Add (" SELECT CANCDATA FROM SFN_DOCUMENTO WHERE HANDLE = :HANDLE ")
              q7.ParamByName("HANDLE").Value = q1.FieldByName("HANDLE").Value
              q7.Active = True

              If q7.FieldByName("CANCDATA").IsNull Then

                q3.Clear
                q3.Add (" SELECT MIN(DATAPAGAMENTO) PGTO  ")
                q3.Add ("   FROM SAM_PAGAMENTO            ")
                q3.Add ("  WHERE DATAFECHAMENTO IS NULL   ")
                q3.Add ("    AND DATAPAGAMENTO >= :HOJE   ")
                q3.Text = SqlConverte(q3.Text, SQLServer)
                q3.ParamByName("HOJE").Value = Date
                q3.Active = True

                If q3.FieldByName("PGTO").IsNull Then
                  MsgBox("Não existe nenhuma data de pagamento aberta")
                  Exit Sub
                End If

                Dim interface1 As Object
                Set interface1 = CreateBennerObject("SFNCANCEL.CANCELAMENTO")
                interface1.CancelDocOnLine(q1.FieldByName("HANDLE").AsInteger, Date, Date, "Alteração de Conta Corrente")
                Set interface1 = Nothing

                '           q2.Clear
                '           q2.Add (" UPDATE SFN_FATURA SET SITUACAO = 'C' WHERE HANDLE = :FATURA ")
                '           q2.Text = SqlConverte(q2.Text, SQLServer)
                '           q2.ParamByName("FATURA").Value = q1.FieldByName("FATURA").AsInteger
                '           q2.ExecSQL

                Dim Interface2 As Object
                Set Interface2 = CreateBennerObject("FINANCEIRO.DOCUMENTO")
                Interface2.CRIAR(q1.FieldByName("CONTAFINANCEIRA").AsInteger, q1.FieldByName("TIPODOCUMENTO").AsInteger, _
                                 Date, q3.FieldByName("PGTO").AsDateTime, _
                                 q1.FieldByName("COMPETENCIA").AsDateTime, q1.FieldByName("NUMERO").AsInteger, _
                                 q1.FieldByName("NOSSONUMERO").AsString, q1.FieldByName("TESOURARIA").AsInteger, _
                                 0, q1.FieldByName("REGRAFINANCEIRA").AsInteger, _
                                 q1.FieldByName("CONTRATO").AsInteger, q1.FieldByName("FAMILIA").AsInteger)
                Set Interface2 = Nothing

                q4.Clear
                q4.Add (" SELECT MAX(HANDLE) HANDLE          ")
                q4.Add ("   FROM SFN_DOCUMENTO               ")
                q4.Add ("  WHERE CONTAFINANCEIRA = :CONTAFIN ")
                q4.ParamByName("CONTAFIN").Value = q.FieldByName("CONTAFINANCEIRA").AsInteger
                q4.Active = True

                q5.Clear
                q5.Add (" INSERT INTO SFN_DOCUMENTO_FATURA ")
                q5.Add (" ( HANDLE,DOCUMENTO,FATURA,SALDO,NATUREZA,VALORJURO,VALORMULTA,VALORCORRECAO,VALORDESCONTO,VALORTOTAL ) ")
                q5.Add (" VALUES ")
                q5.Add (" ( :HANDLE,:DOCUMENTO,:FATURA,:SALDO,:NATUREZA,:VALORJURO,:VALORMULTA,:VALORCORRECAO,:VALORDESCONTO,:VALORTOTAL ) ")
                q5.ParamByName("HANDLE").Value = NewHandle("SFN_DOCUMENTO_FATURA")
                q5.ParamByName("DOCUMENTO").Value = q4.FieldByName("HANDLE").Value
                q5.ParamByName("FATURA").Value = q1.FieldByName("FATURA").Value
                q5.ParamByName("SALDO").Value = q1.FieldByName("SALDO").Value
                q5.ParamByName("NATUREZA").Value = q1.FieldByName("NATUREZA").Value
                q5.ParamByName("VALORJURO").Value = q1.FieldByName("VALORJURO").Value
                q5.ParamByName("VALORMULTA").Value = q1.FieldByName("VALORMULTA").Value
                q5.ParamByName("VALORCORRECAO").Value = q1.FieldByName("VALORCORRECAO").Value
                q5.ParamByName("VALORDESCONTO").Value = q1.FieldByName("VALORDESCONTO").Value
                q5.ParamByName("VALORTOTAL").Value = q1.FieldByName("VALORTOTAL").Value
                q5.ExecSQL

                q6.Clear
                q6.Add (" UPDATE SFN_DOCUMENTO SET VALOR = :VALOR, NATUREZA = :NATU, VALORTOTAL = :VLTOT,  ")
                q6.Add ("        VALORJURO = :JUROS, VALORCORRECAO = :CORRECAO, VALORDESCONTO = :DESCONTO, ")
                q6.Add ("        BANCO = :BANCO, AGENCIA = :AGENCIA, CONTACORRENTE = :CC, DV = :dv         ")
                q6.Add (" WHERE HANDLE = :HANDLE                                                           ")
                q6.ParamByName("HANDLE").Value = q4.FieldByName("HANDLE").Value
                q6.ParamByName("VALOR").Value = q1.FieldByName("VALOR").Value
                q6.ParamByName("NATU").Value = q1.FieldByName("NATU").Value
                q6.ParamByName("VLTOT").Value = q1.FieldByName("VLTOT").Value
                q6.ParamByName("JUROS").Value = q1.FieldByName("JUROS").Value
                q6.ParamByName("CORRECAO").Value = q1.FieldByName("CORRECAO").Value
                q6.ParamByName("DESCONTO").Value = q1.FieldByName("DESCONTO").Value
                q6.ParamByName("BANCO").Value = QContas.FieldByName("BANCO").Value
                q6.ParamByName("AGENCIA").Value = QContas.FieldByName("AGENCIA").Value
                q6.ParamByName("CC").Value = QContas.FieldByName("CONTACORRENTE").Value
                q6.ParamByName("DV").Value = QContas.FieldByName("DV").Value
                q6.ExecSQL


                '           Dim Interface3 As Object
                '           Set Interface3=CreateBennerObject("FINANCEIRO.DOCUMENTO")
                '           Interface3.ATUALIZAR(q4.FieldByName("HANDLE").Value)
                '           Set Interface3=Nothing
              End If
              q1.Next
            Wend

          End If
        End If
      End If
    End If

    QContas.Next
  Wend

  GoTo FINALIZA
ERRO:
  Rollback
  MsgBox Error

FINALIZA:
  Set VERIFICA = Nothing
  Set CONTAFIN = Nothing
  Set ALTERA = Nothing
  Set CODBANCO = Nothing
  Set CODAGENCIA = Nothing
  Set tipodoc = Nothing

  MsgBox "Contas atualizadas !!!"

  '  Dim SQLSELECIONA As Object
  '  Set SQLSELECIONA =NewQuery
  '  Dim RTF As Object
  '  Set RTF =CreateBennerObject("CONTRATO.ContratoInterface")

  '  If Not InTransaction Then
  '    StartTransaction
  '  End If

  '  SQLSELECIONA.Clear
  '  SQLSELECIONA.Add("SELECT HANDLE FROM SAM_CONTRATO ")
  '  SQLSELECIONA.Active =True

  '  While Not SQLSELECIONA.EOF
  '    RTF.INCLUI(CurrentSystem,SQLSELECIONA.FieldByName("HANDLE").AsInteger)
  '
  '    SQLSELECIONA.Next
  '  Wend


  '  If InTransaction Then
  '    Commit
  '  End If

  '  MsgBox "Acabou !!!"

End Sub

Public Sub PROCURA_OnClick()


  Dim Obj As Object
  Set Obj = CreateBennerObject("procura.procurar")
  Obj.Inicializar(CurrentSystem)
  Obj.Exec(CurrentSystem, "SAM_TEXTOPADRAO", "CODIGO|RESUMO", 1, "Código|Resumo", "", "Texto padrão", True, "")
  Set Obj = Nothing


End Sub

Public Sub SAMCONSULTA_OnClick()
  Dim SQL As Object
  Dim UPD As Object
  Dim QCF As Object

  Set SQL = NewQuery
  Set UPD = NewQuery
  Set QCF = NewQuery

  UPD.Clear
  UPD.Add("UPDATE SFN_CONTAFIN SET CLASSECONTABIL=:CLASSECONTABIL")
  UPD.Add("WHERE HANDLE=:HANDLE")

  SQL.Clear
  SQL.Active = False
  SQL.Add("SELECT DISTINCT C.HANDLE CONTRATO, C.LOCALFATURAMENTO,")
  SQL.Add("       C.PESSOA RESPONSAVELCONTRATO, F.TABRESPONSAVEL,")
  SQL.Add("       F.TITULARRESPONSAVEL, F.PESSOARESPONSAVEL")
  SQL.Add("FROM SAM_FAMILIA F, SAM_CONTRATO C")
  SQL.Add("WHERE C.HANDLE=5 AND C.HANDLE=F.CONTRATO")
  SQL.Active = True

  StartTransaction

  While Not SQL.EOF
    QCF.Clear
    QCF.Active = False
    QCF.Add("SELECT HANDLE FROM SFN_CONTAFIN")

    If SQL.FieldByName("LOCALFATURAMENTO").AsString = "C" Then
      QCF.Add("WHERE PESSOA = :RESPONSAVELCONTRATO")
      QCF.ParamByName("RESPONSAVELCONTRATO").AsInteger = SQL.FieldByName("RESPONSAVELCONTRATO").AsInteger
    ElseIf SQL.FieldByName("TABRESPONSAVEL").AsInteger = 1 Then
      QCF.Add("WHERE BENEFICIARIO = :BENEFICIARIO")
      QCF.ParamByName("BENEFICIARIO").AsInteger = SQL.FieldByName("TITULARRESPONSAVEL").AsInteger
    Else
      QCF.Add("WHERE PESSOA = :PESSOA")
      QCF.ParamByName("PESSOA").AsInteger = SQL.FieldByName("PESSOARESPONSAVEL").AsInteger
    End If

    QCF.Active = True

    If Not QCF.EOF Then
      UPD.Active = False
      UPD.ParamByName("CLASSECONTABIL").AsInteger = 20914
      UPD.ParamByName("HANDLE").AsInteger = QCF.FieldByName("HANDLE").AsInteger
      UPD.ExecSQL
    End If
    SQL.Next
  Wend

  Commit

  '=============================================================================================

  SQL.Clear
  SQL.Active = False
  SQL.Add("SELECT DISTINCT C.HANDLE CONTRATO, C.LOCALFATURAMENTO,")
  SQL.Add("       C.PESSOA RESPONSAVELCONTRATO, F.TABRESPONSAVEL,")
  SQL.Add("       F.TITULARRESPONSAVEL, F.PESSOARESPONSAVEL")
  SQL.Add("FROM SAM_FAMILIA F, SAM_CONTRATO C")
  SQL.Add("WHERE C.HANDLE IN (1,2,3,4,6,7,71,72,73,77,595,596,597,598)")
  SQL.Add("      AND C.HANDLE=F.CONTRATO")
  SQL.Active = True

  StartTransaction

  While Not SQL.EOF
    QCF.Clear
    QCF.Active = False
    QCF.Add("SELECT HANDLE FROM SFN_CONTAFIN")

    If SQL.FieldByName("LOCALFATURAMENTO").AsString = "C" Then
      QCF.Add("WHERE PESSOA = :RESPONSAVELCONTRATO")
      QCF.ParamByName("RESPONSAVELCONTRATO").AsInteger = SQL.FieldByName("RESPONSAVELCONTRATO").AsInteger
    ElseIf SQL.FieldByName("TABRESPONSAVEL").AsInteger = 1 Then
      QCF.Add("WHERE BENEFICIARIO = :BENEFICIARIO")
      QCF.ParamByName("BENEFICIARIO").AsInteger = SQL.FieldByName("TITULARRESPONSAVEL").AsInteger
    Else
      QCF.Add("WHERE PESSOA = :PESSOA")
      QCF.ParamByName("PESSOA").AsInteger = SQL.FieldByName("PESSOARESPONSAVEL").AsInteger
    End If

    QCF.Active = True

    If Not QCF.EOF Then
      UPD.Active = False
      UPD.ParamByName("CLASSECONTABIL").AsInteger = 16693
      UPD.ParamByName("HANDLE").AsInteger = QCF.FieldByName("HANDLE").AsInteger
      UPD.ExecSQL
    End If
    SQL.Next
  Wend

  Commit

  '=============================================================================================

  SQL.Clear
  SQL.Active = False
  SQL.Add("SELECT DISTINCT C.HANDLE CONTRATO, C.LOCALFATURAMENTO,")
  SQL.Add("       C.PESSOA RESPONSAVELCONTRATO, F.TABRESPONSAVEL,")
  SQL.Add("       F.TITULARRESPONSAVEL, F.PESSOARESPONSAVEL")
  SQL.Add("FROM SAM_FAMILIA F, SAM_CONTRATO C")
  SQL.Add("WHERE C.CONTRATO > 9000 AND C.HANDLE=F.CONTRATO")
  SQL.Active = True

  StartTransaction

  While Not SQL.EOF
    QCF.Clear
    QCF.Active = False
    QCF.Add("SELECT HANDLE FROM SFN_CONTAFIN")

    If SQL.FieldByName("LOCALFATURAMENTO").AsString = "C" Then
      QCF.Add("WHERE PESSOA = :RESPONSAVELCONTRATO")
      QCF.ParamByName("RESPONSAVELCONTRATO").AsInteger = SQL.FieldByName("RESPONSAVELCONTRATO").AsInteger
    ElseIf SQL.FieldByName("TABRESPONSAVEL").AsInteger = 1 Then
      QCF.Add("WHERE BENEFICIARIO = :BENEFICIARIO")
      QCF.ParamByName("BENEFICIARIO").AsInteger = SQL.FieldByName("TITULARRESPONSAVEL").AsInteger
    Else
      QCF.Add("WHERE PESSOA = :PESSOA")
      QCF.ParamByName("PESSOA").AsInteger = SQL.FieldByName("PESSOARESPONSAVEL").AsInteger
    End If

    QCF.Active = True

    If Not QCF.EOF Then
      UPD.Active = False
      UPD.ParamByName("CLASSECONTABIL").AsInteger = 16741
      UPD.ParamByName("HANDLE").AsInteger = QCF.FieldByName("HANDLE").AsInteger
      UPD.ExecSQL
    End If
    SQL.Next
  Wend

  Commit

  '=============================================================================================

  SQL.Clear
  SQL.Active = False
  SQL.Add("SELECT HANDLE FROM SFN_CONTAFIN")
  SQL.Add("WHERE CLASSECONTABIL IS NULL")
  SQL.Active = True

  StartTransaction

  While Not SQL.EOF
    UPD.Active = False
    UPD.ParamByName("CLASSECONTABIL").AsInteger = 16690
    UPD.ParamByName("HANDLE").AsInteger = SQL.FieldByName("HANDLE").AsInteger
    UPD.ExecSQL

    SQL.Next
  Wend

  Commit

End Sub


Public Function ProcuraPrestador(CPF_Nome As String, Sol_Exe_Rec_Todos As String, TextoPrestador As String ) As Long

  Dim Interface As Object
  Dim vCriterio As String
  Dim qparametros As Object

  Set qparametros = NewQuery
  Dim vCampos As String
  Dim vColunas As String

  qparametros.Add("SELECT UTILIZARCONSULTACENTRAL FROM SAM_PARAMETROSATENDIMENTO")
  qparametros.Active = True
  If qparametros.FieldByName("UTILIZARCONSULTACENTRAL").AsString = "S" Then

    Set Interface = CreateBennerObject("CA009.ConsultaPrestador")

    If CPF_Nome = "C" Then

      Dim vAux As String
      Dim vPosicao As Integer

      vAux = Left(TextoPrestador, 1)

      If Len(vAux) > 0 Then
        If vAux = "0" Or _
                   vAux = "1" Or _
                   vAux = "2" Or _
                   vAux = "3" Or _
                   vAux = "4" Or _
                   vAux = "5" Or _
                   vAux = "6" Or _
                   vAux = "7" Or _
                   vAux = "8" Or _
                   vAux = "9" Then
          vPosicao = 0
        Else
          vPosicao = 1
        End If
      End If



      If (Sol_Exe_Rec_Todos = "L") Then
        ProcuraPrestador = Interface.Filtro(CurrentSystem, vPosicao, TextoPrestador, "L")
      Else
        If (Sol_Exe_Rec_Todos = "S") Then
          ProcuraPrestador = Interface.Filtro(CurrentSystem, vPosicao, TextoPrestador , "S")
        Else
          If (Sol_Exe_Rec_Todos = "E") Then
            ProcuraPrestador = Interface.Filtro(CurrentSystem, vPosicao, TextoPrestador , "E")
          Else
            If (Sol_Exe_Rec_Todos = "R") Then
              ProcuraPrestador = Interface.Filtro(CurrentSystem, vPosicao, TextoPrestador , "R")
            Else
              If (Sol_Exe_Rec_Todos = "T") Then
                ProcuraPrestador = Interface.Filtro(CurrentSystem, vPosicao, TextoPrestador , "T")

              End If
            End If
          End If
        End If
      End If
    Else
      If (Sol_Exe_Rec_Todos = "L") Then
        ProcuraPrestador = Interface.Filtro(CurrentSystem, 1, TextoPrestador , "L")
      Else
        If (Sol_Exe_Rec_Todos = "S") Then
          ProcuraPrestador = Interface.Filtro(CurrentSystem, 1, TextoPrestador , "S")
        Else
          If (Sol_Exe_Rec_Todos = "E") Then
            ProcuraPrestador = Interface.Filtro(CurrentSystem, 1, TextoPrestador , "E")
          Else
            If (Sol_Exe_Rec_Todos = "R") Then
              ProcuraPrestador = Interface.Filtro(CurrentSystem, 1, TextoPrestador , "R")
            Else
              If (Sol_Exe_Rec_Todos = "T") Then
                ProcuraPrestador = Interface.Filtro(CurrentSystem, 1, TextoPrestador , "T")
              End If
            End If
          End If
        End If
      End If

    End If


    Set Interface = Nothing
  End If

  If qparametros.FieldByName("UTILIZARCONSULTACENTRAL").AsString = "N" Then



    Set Interface = CreateBennerObject("Procura.Procurar")


    vColunas = "SAM_PRESTADOR.PRESTADOR"
    vColunas = vColunas + "|SAM_PRESTADOR.Z_NOME"
    vColunas = vColunas + "|SAM_PRESTADOR.INSCRICAOCR"
    vColunas = vColunas + "|SAM_PRESTADOR.DATACREDENCIAMENTO"
    vColunas = vColunas + "|SAM_PRESTADOR.DATADESCREDENCIAMENTO"
    vColunas = vColunas + "|SAM_PRESTADOR.SOLICITANTE"
    vColunas = vColunas + "|SAM_PRESTADOR.EXECUTOR"
    vColunas = vColunas + "|SAM_PRESTADOR.RECEBEDOR"
    vColunas = vColunas + "|ESTADOS.NOME NOMEESTADO"
    vColunas = vColunas + "|MUNICIPIOS.NOME NOMEMUNICIPIO"

    vCriterio = "(1=1)"

    If (Sol_Exe_Rec_Todos = "L") Then
      vCriterio = vCriterio + " AND SAM_PRESTADOR.LOCALEXECUCAO = 'S' "
    Else
      If (Sol_Exe_Rec_Todos = "S") Then
        vCriterio = vCriterio + " AND SAM_PRESTADOR.SOLICITANTE = 'S' "
      Else
        If (Sol_Exe_Rec_Todos = "E") Then
          vCriterio = vCriterio + " AND SAM_PRESTADOR.EXECUTOR = 'S' "
        Else
          If (Sol_Exe_Rec_Todos = "R") Then
            vCriterio = vCriterio + " AND SAM_PRESTADOR.RECEBEDOR = 'S'  "
          End If
        End If
      End If
    End If


    vCampos = "Prestador|Nome do Prestador|Nr.Conselho|Credenciam.|Descredenc|Sol|Exe|Rec|Estado|Município"

    If CPF_Nome = "C" Then
      ProcuraPrestador = Interface.Exec(CurrentSystem, "SAM_PRESTADOR|*SAM_CONSELHO[SAM_CONSELHO.HANDLE =SAM_PRESTADOR.CONSELHOREGIONAL]|*ESTADOS[ESTADOS.HANDLE =SAM_PRESTADOR.ESTADOPAGAMENTO]|*MUNICIPIOS[MUNICIPIOS.HANDLE =SAM_PRESTADOR.MUNICIPIOPAGAMENTO]", vColunas, 1, vCampos, vCriterio, "Prestadores", False, TextoPrestador, "CA005.ConsultaPrestador")
    Else
      ProcuraPrestador = Interface.Exec(CurrentSystem, "SAM_PRESTADOR|*SAM_CONSELHO[SAM_CONSELHO.HANDLE =SAM_PRESTADOR.CONSELHOREGIONAL]|*ESTADOS[ESTADOS.HANDLE =SAM_PRESTADOR.ESTADOPAGAMENTO]|*MUNICIPIOS[MUNICIPIOS.HANDLE =SAM_PRESTADOR.MUNICIPIOPAGAMENTO]", vColunas, 2, vCampos, vCriterio, "Prestadores", False, TextoPrestador, "CA005.ConsultaPrestador")
    End If

    Set Interface = Nothing

  End If
End Function

Public Sub TABLE_ExternalValidate(CanContinue As Boolean, ByVal Param As String)

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "TESTE" Then
		CA001_OnClick
	End If
End Sub
