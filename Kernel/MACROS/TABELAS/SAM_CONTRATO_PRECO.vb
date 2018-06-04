'HASH: 7F2DDFB6D63E3EF840465B7984DF2DBC
'Macro: SAM_CONTRATO_PRECO
Option Explicit

'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Public Sub BOTAOGERAPRESTADOR_OnClick()
  Dim sql As Object
  Dim vMenosPrestadores As String
  Dim InterfacePrestador As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vMostrar As String
  Dim vColunas As String
  Dim vCriterio As String

  If CurrentQuery.State <> 1 Then
    bsShowMessage("O registro não pode estar em edição.", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("DATAFINAL").AsString <> "" Then
    Exit Sub
  End If

  vColunas = "NOME"
  vCampos = "NOME"
  vMostrar = "S"

  '****************************************       Verifica se existe tabela de preço         ****************************************
  Set sql = NewQuery
  sql.Clear
  sql.Add("SELECT COUNT(*) QTDE FROM SAM_CONTRATO_PRECO WHERE CONTRATO = " + CurrentQuery.FieldByName("CONTRATO").AsString)
  If Not CurrentQuery.FieldByName("TABELAPRECO").IsNull Then
    sql.Add("   AND TABELAPRECO  = " + CurrentQuery.FieldByName("TABELAPRECO").AsString)
  End If
  sql.Add("   AND PLANO = " + CurrentQuery.FieldByName("PLANO").AsString)'Anderson sms 21638
  sql.Active = True

  If sql.FieldByName("QTDE").AsInteger >1 Then 'existe já esta tabela cadastrada
    'verifica faixa de evento na tabela
    sql.Clear
    sql.Add("SELECT COUNT(*) QTDE FROM SAM_CONTRATO_PRECO ")
    sql.Add(" WHERE CONTRATO = " + CurrentQuery.FieldByName("CONTRATO").AsString)
    sql.Add("   AND TABELAPRECO  = " + CurrentQuery.FieldByName("TABELAPRECO").AsString)
    sql.Add("   And (REPLACE('" + CurrentQuery.FieldByName("ESTRUTURAFINAL").AsString + "','.','') BETWEEN REPLACE(ESTRUTURAINICIAL,'.','') AND REPLACE(ESTRUTURAFINAL,'.','')      ")
    sql.Add("   OR   REPLACE('" + CurrentQuery.FieldByName("ESTRUTURAINICIAL").AsString + "','.','') BETWEEN REPLACE(ESTRUTURAINICIAL,'.','') AND REPLACE(ESTRUTURAFINAL,'.','')     )")
    sql.Active = True

    If sql.FieldByName("QTDE").AsInteger >1 Then 'existe evento lançado na mesma tabela
      'Evento já lançado
      sql.Clear
      sql.Add("SELECT PRESTADOR FROM SAM_CONTRATO_PRECO_PRESTADOR                              ")
      sql.Add(" WHERE CONTRATOPRECO IN (SELECT HANDLE FROM SAM_CONTRATO_PRECO                  ")
      sql.Add("                          WHERE CONTRATO = " + CurrentQuery.FieldByName("CONTRATO").AsString + " AND PLANO = " + CurrentQuery.FieldByName("PLANO").AsString)'Anderson sms 21638
      sql.Add("   And (REPLACE('" + CurrentQuery.FieldByName("ESTRUTURAFINAL").AsString + "','.','') BETWEEN REPLACE(ESTRUTURAINICIAL,'.','') AND REPLACE(ESTRUTURAFINAL,'.','')      ")
      sql.Add("   OR   REPLACE('" + CurrentQuery.FieldByName("ESTRUTURAINICIAL").AsString + "','.','') BETWEEN REPLACE(ESTRUTURAINICIAL,'.','') AND REPLACE(ESTRUTURAFINAL,'.','')     )")
      sql.Add("   AND TABELAPRECO = " + CurrentQuery.FieldByName("TABELAPRECO").AsString)
      sql.Add(" )")
      sql.Active = True
      vMenosPrestadores = ""
      While Not sql.EOF
        If vMenosPrestadores = "" Then
          vMenosPrestadores = sql.FieldByName("PRESTADOR").AsString
        Else
          vMenosPrestadores = vMenosPrestadores + "," + sql.FieldByName("PRESTADOR").AsString
        End If
        sql.Next
      Wend
      Set InterfacePrestador = CreateBennerObject("Procura.Procurar")
      'Parâmetros(pTabela       ,pColuna   ,pCampos ,pTabAssoc                     ,pCampoAssoc       ,pHandleAssoc                             ,pCampoAssoc2,pTitulo: WideString),pSqlEspecial    ,pMostrar  pMunicipio
      InterfacePrestador.sELECIONA(CurrentSystem, "SAM_PRESTADOR", vColunas, vCampos, "SAM_CONTRATO_PRECO_PRESTADOR", "CONTRATOPRECO", RecordHandleOfTable("SAM_CONTRATO_PRECO"), "PRESTADOR", "Lista de Prestadores", vMenosPrestadores, vMostrar, "S", "C")
      Set InterfacePrestador = Nothing
      sql.Clear
      sql.Add("SELECT COUNT(*) NREC FROM SAM_CONTRATO_PRECO P WHERE P.CONTRATO=" + CurrentQuery.FieldByName("CONTRATO").AsString + " AND PLANO = " + CurrentQuery.FieldByName("PLANO").AsString)'Anderson sms 21638
      sql.Add("   And Not EXISTS (Select HANDLE FROM SAM_CONTRATO_PRECO_PRESTADOR PP WHERE PP.CONTRATOPRECO=P.HANDLE)")
      sql.Add("   AND TABELAPRECO=" + CurrentQuery.FieldByName("TABELAPRECO").AsString)
      sql.Add("   AND DATAFINAL IS NULL")
      sql.Add("   And (REPLACE('" + CurrentQuery.FieldByName("ESTRUTURAFINAL").AsString + "','.','') BETWEEN REPLACE(ESTRUTURAINICIAL,'.','') AND REPLACE(ESTRUTURAFINAL,'.','')      ")
      sql.Add("   OR   REPLACE('" + CurrentQuery.FieldByName("ESTRUTURAINICIAL").AsString + "','.','') BETWEEN REPLACE(ESTRUTURAINICIAL,'.','') AND REPLACE(ESTRUTURAFINAL,'.','')     )")
      sql.Active = True
      If sql.FieldByName("NREC").AsInteger >1 Then 'VERIFICA SE JÁ EXISTE ALGUMA TABELA DE PRECO SEM PRESTADORES
        CurrentQuery.Delete 'Apaga
      End If
    Else
      'Evento ainda não lançado nesta tabela
      Set sql = NewQuery
      sql.Clear
      sql.Add("SELECT PRESTADOR FROM SAM_CONTRATO_PRECO_PRESTADOR                              ")
      sql.Add(" WHERE CONTRATOPRECO IN (SELECT HANDLE FROM SAM_CONTRATO_PRECO                  ")
      sql.Add("                          WHERE CONTRATO = " + CurrentQuery.FieldByName("CONTRATO").AsString + " AND PLANO = " + CurrentQuery.FieldByName("PLANO").AsString)'Anderson sms 21638
      sql.Add("   And (REPLACE('" + CurrentQuery.FieldByName("ESTRUTURAFINAL").AsString + "','.','') BETWEEN REPLACE(ESTRUTURAINICIAL,'.','') AND REPLACE(ESTRUTURAFINAL,'.','')      ")
      sql.Add("   OR   REPLACE('" + CurrentQuery.FieldByName("ESTRUTURAINICIAL").AsString + "','.','') BETWEEN REPLACE(ESTRUTURAINICIAL,'.','') AND REPLACE(ESTRUTURAFINAL,'.','')     )")
      sql.Add(")")
      sql.Active = True
      vMenosPrestadores = ""
      While Not sql.EOF
        If vMenosPrestadores = "" Then
          vMenosPrestadores = sql.FieldByName("PRESTADOR").AsString
        Else
          vMenosPrestadores = vMenosPrestadores + "," + sql.FieldByName("PRESTADOR").AsString
        End If
        sql.Next
      Wend
      Set InterfacePrestador = CreateBennerObject("Procura.Procurar")
      'Parâmetros(pTabela       ,pColuna   ,pCampos ,pTabAssoc                     ,pCampoAssoc       ,pHandleAssoc                             ,pCampoAssoc2,pTitulo: WideString),pSqlEspecial    ,pMostrar  pMunicipio
      InterfacePrestador.sELECIONA(CurrentSystem, "SAM_PRESTADOR", vColunas, vCampos, "SAM_CONTRATO_PRECO_PRESTADOR", "CONTRATOPRECO", RecordHandleOfTable("SAM_CONTRATO_PRECO"), "PRESTADOR", "Lista de Prestadores", vMenosPrestadores, vMostrar, "S", "C")
      Set InterfacePrestador = Nothing
    End If
  Else
    sql.Clear
    sql.Add("SELECT COUNT(*) QTDE FROM SAM_CONTRATO_PRECO WHERE CONTRATO = " + CurrentQuery.FieldByName("CONTRATO").AsString + " AND PLANO = " + CurrentQuery.FieldByName("PLANO").AsString)'Anderson sms 21638
    If Not CurrentQuery.FieldByName("TABELAPRECO").IsNull Then
      sql.Add("   AND TABELAPRECO  <> " + CurrentQuery.FieldByName("TABELAPRECO").AsString)
    End If
    sql.Add("   And (REPLACE('" + CurrentQuery.FieldByName("ESTRUTURAFINAL").AsString + "','.','') BETWEEN REPLACE(ESTRUTURAINICIAL,'.','') AND REPLACE(ESTRUTURAFINAL,'.','')      ")
    sql.Add("   OR   REPLACE('" + CurrentQuery.FieldByName("ESTRUTURAINICIAL").AsString + "','.','') BETWEEN REPLACE(ESTRUTURAINICIAL,'.','') AND REPLACE(ESTRUTURAFINAL,'.','')     )")
    sql.Active = True
    If sql.FieldByName("QTDE").AsInteger >0 Then
      'evento já lançado em outra tabela
      sql.Clear
      sql.Add("SELECT P.PRESTADOR FROM SAM_CONTRATO_PRECO_PRESTADOR P                            ")
      sql.Add(" WHERE P.CONTRATOPRECO IN (SELECT HANDLE FROM SAM_CONTRATO_PRECO                  ")
      sql.Add("                          WHERE CONTRATO = " + CurrentQuery.FieldByName("CONTRATO").AsString + " AND PLANO = " + CurrentQuery.FieldByName("PLANO").AsString)'Anderson sms 21638
      sql.Add("   And (REPLACE('" + CurrentQuery.FieldByName("ESTRUTURAFINAL").AsString + "','.','') BETWEEN REPLACE(ESTRUTURAINICIAL,'.','') AND REPLACE(ESTRUTURAFINAL,'.','')      ")
      sql.Add("   OR   REPLACE('" + CurrentQuery.FieldByName("ESTRUTURAINICIAL").AsString + "','.','') BETWEEN REPLACE(ESTRUTURAINICIAL,'.','') AND REPLACE(ESTRUTURAFINAL,'.','')     )")
      sql.Add("                            AND PRESTADOR      = P.PRESTADOR                      ")
      sql.Add("                            AND HANDLE         <> :TABELAPRECO)                    ")
      sql.ParamByName("TABELAPRECO").Value = CurrentQuery.FieldByName("TABELAPRECO").AsInteger
      sql.Active = True
      vMenosPrestadores = ""
      While Not sql.EOF
        If vMenosPrestadores = "" Then
          vMenosPrestadores = sql.FieldByName("PRESTADOR").AsString
        Else
          vMenosPrestadores = vMenosPrestadores + "," + sql.FieldByName("PRESTADOR").AsString
        End If
        sql.Next
      Wend
      Set InterfacePrestador = CreateBennerObject("Procura.Procurar")
      'Parâmetros(pTabela       ,pColuna   ,pCampos ,pTabAssoc                     ,pCampoAssoc       ,pHandleAssoc                             ,pCampoAssoc2,pTitulo: WideString),pSqlEspecial    ,pMostrar  pMunicipio
      InterfacePrestador.sELECIONA(CurrentSystem, "SAM_PRESTADOR", vColunas, vCampos, "SAM_CONTRATO_PRECO_PRESTADOR", "CONTRATOPRECO", RecordHandleOfTable("SAM_CONTRATO_PRECO"), "PRESTADOR", "Lista de Prestadores", vMenosPrestadores, vMostrar, "S", "C")
      Set InterfacePrestador = Nothing
      sql.Clear
      sql.Add("SELECT COUNT(*) NREC FROM SAM_CONTRATO_PRECO P WHERE P.CONTRATO=" + CurrentQuery.FieldByName("CONTRATO").AsString)
      sql.Add("   And Not EXISTS (Select HANDLE FROM SAM_CONTRATO_PRECO_PRESTADOR PP WHERE PP.CONTRATOPRECO=P.HANDLE)")
      sql.Add("   AND DATAFINAL IS NULL")
      sql.Add("   And (REPLACE('" + CurrentQuery.FieldByName("ESTRUTURAFINAL").AsString + "','.','') BETWEEN REPLACE(ESTRUTURAINICIAL,'.','') AND REPLACE(ESTRUTURAFINAL,'.','')      ")
      sql.Add("   OR   REPLACE('" + CurrentQuery.FieldByName("ESTRUTURAINICIAL").AsString + "','.','') BETWEEN REPLACE(ESTRUTURAINICIAL,'.','') AND REPLACE(ESTRUTURAFINAL,'.','')     )")
      sql.Active = True
      If sql.FieldByName("NREC").AsInteger >1 Then 'VERIFICA SE JÁ EXISTE ALGUMA TABELA DE PRECO SEM PRESTADORES
        CurrentQuery.Delete 'Apaga
      End If
    Else
      'evento ainda não lançado em outra tabela
      sql.Clear
      sql.Add("SELECT PRESTADOR FROM SAM_CONTRATO_PRECO_PRESTADOR                              ")
      sql.Add(" WHERE CONTRATOPRECO IN (SELECT HANDLE FROM SAM_CONTRATO_PRECO WHERE CONTRATO = " + CurrentQuery.FieldByName("CONTRATO").AsString + " AND PLANO = " + CurrentQuery.FieldByName("PLANO").AsString + " AND TABELAPRECO = " + CurrentQuery.FieldByName("TABELAPRECO").AsString + ")")
      sql.Active = True
      vMenosPrestadores = ""
      While Not sql.EOF
        If vMenosPrestadores = "" Then
          vMenosPrestadores = sql.FieldByName("PRESTADOR").AsString
        Else
          vMenosPrestadores = vMenosPrestadores + "," + sql.FieldByName("PRESTADOR").AsString
        End If
        sql.Next
      Wend
      Set InterfacePrestador = CreateBennerObject("Procura.Procurar")
      'Parâmetros(pTabela       ,pColuna   ,pCampos ,pTabAssoc                     ,pCampoAssoc       ,pHandleAssoc                             ,pCampoAssoc2,pTitulo: WideString),pSqlEspecial    ,pMostrar  pMunicipio
      InterfacePrestador.sELECIONA(CurrentSystem, "SAM_PRESTADOR", vColunas, vCampos, "SAM_CONTRATO_PRECO_PRESTADOR", "CONTRATOPRECO", RecordHandleOfTable("SAM_CONTRATO_PRECO"), "PRESTADOR", "Lista de Prestadores", vMenosPrestadores, vMostrar, "S", "C")
      Set InterfacePrestador = Nothing
    End If
  End If
  Set sql = Nothing
  Set InterfacePrestador = Nothing
End Sub

Public Sub EVENTOINICIAL_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTOINICIAL.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOINICIAL").Value = vHandle
  End If
End Sub

Public Sub EVENTOFINAL_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTOFINAL.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOFINAL").Value = vHandle
  End If
End Sub

Public Function CheckEventosFx As Boolean
  CheckEventosFx = True
  If Not CurrentQuery.FieldByName("EVENTOINICIAL").IsNull Then
    If CurrentQuery.FieldByName("EVENTOFINAL").IsNull Then
      CurrentQuery.FieldByName("EVENTOFINAL").Value = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
    Else
      If CurrentQuery.FieldByName("EVENTOINICIAL").Value <>CurrentQuery.FieldByName("EVENTOFINAL").Value Then
        Dim SQLI, SQLF As Object
        Set SQLI = NewQuery
        SQLI.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :HEVENTOI")
        SQLI.ParamByName("HEVENTOI").Value = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
        SQLI.Active = True

        Set SQLF = NewQuery
        SQLF.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :HEVENTOF")
        SQLF.ParamByName("HEVENTOF").Value = CurrentQuery.FieldByName("EVENTOFINAL").AsInteger
        SQLF.Active = True

        If SQLF.FieldByName("ESTRUTURA").Value <SQLI.FieldByName("ESTRUTURA").Value Then
          bsShowMessage("Evento final não pode ser menor que o evento inicial!", "I")
          EVENTOFINAL.SetFocus
          CheckEventosFx = False
        End If
        Set SQLI = Nothing
        Set SQLF = Nothing
      End If
    End If
  End If
End Function

Public Sub TABLE_AfterPost()
  TABLE_AfterScroll
  If (VisibleMode) Then
    BOTAOGERAPRESTADOR_OnClick
  End If
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DATAFINAL.ReadOnly = False
  Else
    DATAFINAL.ReadOnly = True
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  If WebMode Then
  	PLANO.WebLocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CAMPO(CONTRATO))"
  ElseIf VisibleMode Then
 	PLANO.LocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CONTRATO)"
  End If


 If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    CanContinue = False
    bsShowMessage("Registro finalizado não pode ser alterado!", "E")
    Exit Sub
  End If

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If WebMode Then
  	PLANO.WebLocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CAMPO(CONTRATO))"
  ElseIf VisibleMode Then
 	PLANO.LocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CONTRATO)"
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  ' Atribuir ESTRUTURAINICIAL E FINAL
  Dim SQLTGE, SQLMASC As Object
  Dim Estrutura As String

  Dim EstruturaI As String
  Dim EstruturaF As String
  Dim Interface As Object
  Dim Condicao As String
  Dim Linha As String

  ' Atribuir ESTRUTURAINICIAL
  Set SQLTGE = NewQuery
  SQLTGE.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :HEVENTO")
  SQLTGE.ParamByName("HEVENTO").Value = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
  SQLTGE.Active = True
  CurrentQuery.FieldByName("ESTRUTURAINICIAL").Value = SQLTGE.FieldByName("ESTRUTURA").Value

  ' Atribuir ESTRUTURAFINAL
  SQLTGE.Active = False
  SQLTGE.ParamByName("HEVENTO").Value = CurrentQuery.FieldByName("EVENTOFINAL").AsInteger
  SQLTGE.Active = True
  Estrutura = SQLTGE.FieldByName("ESTRUTURA").Value
  SQLTGE.Active = False
  Set SQLTGE = Nothing

  ' Completar ESTRUTURAFinal com 99999
  Set SQLMASC = NewQuery
  SQLMASC.Add("SELECT M.MASCARA MASCARA FROM Z_TABELAS T, Z_MASCARAS M")
  SQLMASC.Add("WHERE T.NOME = 'SAM_TGE' AND M.TABELA = T.HANDLE")
  SQLMASC.Active = True
  Estrutura = Estrutura + Mid(SQLMASC.FieldByName("MASCARA").AsString, Len(Estrutura) + 1)
  CurrentQuery.FieldByName("ESTRUTURAFINAL").Value = Estrutura
  SQLMASC.Active = False
  Set SQLMASC = Nothing
  CanContinue = CheckEventosFx

  If CanContinue And VisibleMode Then
    ' Checar Vigencia

    EstruturaI = CurrentQuery.FieldByName("ESTRUTURAINICIAL").AsString
    EstruturaF = CurrentQuery.FieldByName("ESTRUTURAFINAL").AsString
    Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
    Condicao = " CONTRATO = " + CurrentQuery.FieldByName("CONTRATO").AsString

    Linha = Interface.EventoFx(CurrentSystem, "SAM_CONTRATO_PRECO", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, EstruturaI, EstruturaF, Condicao)

    If Linha = "" Then
      CanContinue = True
    Else
      Dim sql As Object
      Set sql = NewQuery
      If CurrentQuery.State = 3 Then
        sql.Clear
        sql.Add("SELECT COUNT(HANDLE) QTDE FROM SAM_CONTRATO_PRECO                 ")
        sql.Add(" WHERE CONTRATO = :CONTRATO                                       ")
        sql.ParamByName("CONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
        sql.Add("   AND TABELAUS = :TABELAUS                                       ")
        sql.ParamByName("TABELAUS").Value = CurrentQuery.FieldByName("TABELAUS").AsInteger
        sql.Active = True
      End If
    End If
    Set Interface = Nothing
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOGERAPRESTADOR" Then
		BOTAOGERAPRESTADOR_OnClick
	End If
End Sub
