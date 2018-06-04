'HASH: 575C7A52C308519CBE6188B7FE0F7525
'Macro: SAM_PLANO_PRECO
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

  If CurrentQuery.FieldByName("DATAFINAL").AsString <>"" Then
    Exit Sub
  End If


  If CurrentQuery.State <> 1 Then
    bsShowMessage("O registro não pode estar em inserção ou edição!", "I")
    Exit Sub
  End If


  vColunas = "NOME"
  vCampos = "NOME"
  vMostrar = "S"
  '****************************************       Verifica se existe tabela de preço         ****************************************
  Set sql = NewQuery
  sql.Clear
  sql.Add("SELECT COUNT(*) QTDE FROM SAM_PLANO_PRECO WHERE PLANO = " + CurrentQuery.FieldByName("PLANO").AsString)
  sql.Add("   AND TABELAPRECO  = " + CurrentQuery.FieldByName("TABELAPRECO").AsString)
  sql.Active = True
  '**********************************************************************************************************************************
  If sql.FieldByName("QTDE").AsInteger >1 Then 'existe já esta tabela cadastrada
    'MsgBox("ESTA TABELA JA ESTÁ CADASTRADA")
    'verifica faixa de evento na tabela
    sql.Clear
    sql.Add("SELECT COUNT(*) QTDE FROM SAM_PLANO_PRECO ")
    sql.Add(" WHERE PLANO = " + CurrentQuery.FieldByName("PLANO").AsString)
    sql.Add("   AND TABELAPRECO  = " + CurrentQuery.FieldByName("TABELAPRECO").AsString)
    sql.Add("   And (REPLACE('" + CurrentQuery.FieldByName("ESTRUTURAFINAL").AsString + "','.','') BETWEEN REPLACE(ESTRUTURAINICIAL,'.','') AND REPLACE(ESTRUTURAFINAL,'.','')      ")
    sql.Add("   OR   REPLACE('" + CurrentQuery.FieldByName("ESTRUTURAINICIAL").AsString + "','.','') BETWEEN REPLACE(ESTRUTURAINICIAL,'.','') AND REPLACE(ESTRUTURAFINAL,'.','')     )")
    sql.Active = True
    '**********************************************************************************************************************************
    If sql.FieldByName("QTDE").AsInteger >1 Then 'existe evento lançado na mesma tabela
      'Evento já lançado
      'MsgBox("Evento já cadastrado nesta tabela, que já foi cadastrada")
      sql.Clear
      sql.Add("SELECT PRESTADOR FROM SAM_PLANO_PRECO_PRESTADOR                              ")
      sql.Add(" WHERE CONTRATOPRECO IN (SELECT HANDLE FROM SAM_PLANO_PRECO                  ")
      sql.Add("                          WHERE PLANO = " + CurrentQuery.FieldByName("PLANO").AsString)
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
      InterfacePrestador.sELECIONA(CurrentSystem, "SAM_PRESTADOR", vColunas, vCampos, "SAM_PLANO_PRECO_PRESTADOR", "CONTRATOPRECO", RecordHandleOfTable("SAM_PLANO_PRECO"), "PRESTADOR", "Lista de Prestadores", vMenosPrestadores, vMostrar, "S", "P")
      Set InterfacePrestador = Nothing
      sql.Clear
      sql.Add("SELECT COUNT(*) NREC FROM SAM_PLANO_PRECO P WHERE P.PLANO=" + CurrentQuery.FieldByName("PLANO").AsString)
      sql.Add("   And Not EXISTS (Select HANDLE FROM SAM_PLANO_PRECO_PRESTADOR PP WHERE PP.CONTRATOPRECO=P.HANDLE)")
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
      'MsgBox("Evento ainda não cadastrado nesta tabela, que já foi cadastrada")
      Set sql = NewQuery
      sql.Clear
      sql.Add("SELECT PRESTADOR FROM SAM_PLANO_PRECO_PRESTADOR                              ")
      sql.Add(" WHERE CONTRATOPRECO IN (SELECT HANDLE FROM SAM_PLANO_PRECO                  ")
      sql.Add("                          WHERE PLANO = " + CurrentQuery.FieldByName("PLANO").AsString)
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
      InterfacePrestador.sELECIONA(CurrentSystem, "SAM_PRESTADOR", vColunas, vCampos, "SAM_PLANO_PRECO_PRESTADOR", "CONTRATOPRECO", RecordHandleOfTable("SAM_PLANO_PRECO"), "PRESTADOR", "Lista de Prestadores", vMenosPrestadores, vMostrar, "S", "P")
      Set InterfacePrestador = Nothing
    End If
  Else
    'MsgBox("ESTA TABELA NÃO TEM CADASTRO AINDA")
    sql.Clear
    sql.Add("SELECT COUNT(*) QTDE FROM SAM_PLANO_PRECO WHERE PLANO = " + CurrentQuery.FieldByName("PLANO").AsString)
    sql.Add("   AND TABELAPRECO  <> " + CurrentQuery.FieldByName("TABELAPRECO").AsString)
    sql.Add("   And (REPLACE('" + CurrentQuery.FieldByName("ESTRUTURAFINAL").AsString + "','.','') BETWEEN REPLACE(ESTRUTURAINICIAL,'.','') AND REPLACE(ESTRUTURAFINAL,'.','')      ")
    sql.Add("   OR   REPLACE('" + CurrentQuery.FieldByName("ESTRUTURAINICIAL").AsString + "','.','') BETWEEN REPLACE(ESTRUTURAINICIAL,'.','') AND REPLACE(ESTRUTURAFINAL,'.','')     )")
    sql.Active = True
    If sql.FieldByName("QTDE").AsInteger >0 Then
      'evento já lançado em outra tabela
      'MsgBox("evento já lançado em outra tabela")
      sql.Clear
      sql.Add("SELECT P.PRESTADOR FROM SAM_PLANO_PRECO_PRESTADOR P                            ")
      sql.Add(" WHERE P.CONTRATOPRECO IN (SELECT HANDLE FROM SAM_PLANO_PRECO                  ")
      sql.Add("                          WHERE PLANO = " + CurrentQuery.FieldByName("PLANO").AsString)
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
      InterfacePrestador.sELECIONA(CurrentSystem, "SAM_PRESTADOR", vColunas, vCampos, "SAM_PLANO_PRECO_PRESTADOR", "CONTRATOPRECO", RecordHandleOfTable("SAM_PLANO_PRECO"), "PRESTADOR", "Lista de Prestadores", vMenosPrestadores, vMostrar, "S", "P")
      Set InterfacePrestador = Nothing
      sql.Clear
      sql.Add("SELECT COUNT(*) NREC FROM SAM_PLANO_PRECO P WHERE P.PLANO=" + CurrentQuery.FieldByName("PLANO").AsString)
      sql.Add("   And Not EXISTS (Select HANDLE FROM SAM_PLANO_PRECO_PRESTADOR PP WHERE PP.CONTRATOPRECO=P.HANDLE)")
      'sql.Add("   AND TABELAPRECO<>"+CurrentQuery.FieldByName("TABELAPRECO").AsString)
      sql.Add("   AND DATAFINAL IS NULL")
      sql.Add("   And (REPLACE('" + CurrentQuery.FieldByName("ESTRUTURAFINAL").AsString + "','.','') BETWEEN REPLACE(ESTRUTURAINICIAL,'.','') AND REPLACE(ESTRUTURAFINAL,'.','')      ")
      sql.Add("   OR   REPLACE('" + CurrentQuery.FieldByName("ESTRUTURAINICIAL").AsString + "','.','') BETWEEN REPLACE(ESTRUTURAINICIAL,'.','') AND REPLACE(ESTRUTURAFINAL,'.','')     )")
      sql.Active = True
      If sql.FieldByName("NREC").AsInteger >1 Then 'VERIFICA SE JÁ EXISTE ALGUMA TABELA DE PRECO SEM PRESTADORES
        CurrentQuery.Delete 'Apaga
      End If
    Else
      'evento ainda não lançado em outra tabela
      'MsgBox("evento ainda não lançado em outra tabela")
      sql.Clear
      sql.Add("SELECT PRESTADOR FROM SAM_PLANO_PRECO_PRESTADOR                              ")
      sql.Add(" WHERE CONTRATOPRECO IN (SELECT HANDLE FROM SAM_PLANO_PRECO WHERE PLANO = " + CurrentQuery.FieldByName("PLANO").AsString + " AND TABELAPRECO = " + CurrentQuery.FieldByName("TABELAPRECO").AsString + ")")
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
      InterfacePrestador.sELECIONA(CurrentSystem, "SAM_PRESTADOR", vColunas, vCampos, "SAM_PLANO_PRECO_PRESTADOR", "CONTRATOPRECO", RecordHandleOfTable("SAM_PLANO_PRECO"), "PRESTADOR", "Lista de Prestadores", vMenosPrestadores, vMostrar, "S", "P")
      Set InterfacePrestador = Nothing
    End If
  End If
  Set sql = Nothing
  Set InterfacePrestador = Nothing

End Sub


Public Sub EVENTOINICIAL_OnPopup(ShowPopup As Boolean)
  '  If Len(EVENTO.Text)=0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTOINICIAL.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOINICIAL").Value = vHandle
  End If
  '  End If
End Sub

Public Sub EVENTOFINAL_OnPopup(ShowPopup As Boolean)
  '  If Len(EVENTO.Text)=0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTOFINAL.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOFINAL").Value = vHandle
  End If
  '  End If
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
          bsShowMessage("Evento final não pode ser menor que o evento inicial!", "E")
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
  If VisibleMode Then
  	BOTAOGERAPRESTADOR_OnClick
  End If
End Sub

Public Sub TABLE_AfterScroll()
  'EVENTOINICIAL.AnyLevel=True
  'EVENTOFINAL.AnyLevel=True
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DATAFINAL.ReadOnly = False
  Else
    DATAFINAL.ReadOnly = True
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    CanContinue = False
    bsShowMessage("Registro finalizado não pode ser alterado!", "E")
    Exit Sub
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

  If CanContinue = True Then
    ' Checar Vigencia

    EstruturaI = CurrentQuery.FieldByName("ESTRUTURAINICIAL").AsString
    EstruturaF = CurrentQuery.FieldByName("ESTRUTURAFINAL").AsString

    Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
    Condicao = " PLANO = " + CurrentQuery.FieldByName("PLANO").AsString

    Linha = Interface.EventoFx(CurrentSystem, "SAM_PLANO_PRECO", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, EstruturaI, EstruturaF, Condicao)

    If Linha <> "" Then
      CanContinue = False
      bsShowMessage(Linha, "E")
      Exit Sub
    End If
    Set Interface = Nothing
  End If

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOGERAPRESTADOR" Then
		BOTAOGERAPRESTADOR_OnClick
	End If
End Sub

