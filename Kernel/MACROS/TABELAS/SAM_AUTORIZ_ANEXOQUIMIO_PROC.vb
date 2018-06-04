'HASH: 121067E441DE1792586314FDA70AC86E
'#Uses "*bsShowMessage"
Option Explicit

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  If VisibleMode Then

    If CurrentQuery.FieldByName("CODIGOTABELA").IsNull Then
      bsShowMessage("Selecione o código de tabela antes de selecionar o evento!", "E")
      ShowPopup = False
      Exit Sub
    End If

    Dim interface As Object
    Dim vHandle As Long
    Dim vCampos As String
    Dim vColunas As String
    Dim vCriterio As String

    ShowPopup = False
    Set interface = CreateBennerObject("Procura.Procurar")

    vColunas = "ESTRUTURA|DESCRICAO"

    vCriterio = "INCLUINOTIPOANEXO = 'Q' AND ULTIMONIVEL = 'S' AND INATIVO <> 'S' AND EXISTS(SELECT 1 FROM SAM_TGE_TABELATISS T WHERE T.EVENTO = SAM_TGE.HANDLE AND T.TABELATISS = " + CurrentQuery.FieldByName("CODIGOTABELA").AsString + ")"
    vCampos = "Evento|Descrição"

    vHandle = interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 1, vCampos, vCriterio, "Tabela Geral de Eventos", True, "")

    If vHandle <>0 Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("EVENTO").Value = vHandle
    End If
    Set interface = Nothing

  End If
End Sub

Public Sub TABLE_AfterScroll()
  Dim vSQL As Object
  Set vSQL = NewQuery

  vSQL.Active = False
  vSQL.Clear

  vSQL.Add("SELECT AUTORIZACAO ")
  vSQL.Add("FROM SAM_AUTORIZ_ANEXOQUIMIO ")
  vSQL.Add("WHERE HANDLE = :ANEXOQUIMIO ")
  vSQL.ParamByName("ANEXOQUIMIO").Value = CurrentQuery.FieldByName("ANEXOQUIMIO").AsInteger

  vSQL.Active = True
  
  CODIGOTABELA.WebLocalWhere = "A.VERSAOTISS = "+ CStr(PegarHandleVersaoTISS(vSQL.FieldByName("AUTORIZACAO").AsInteger))
  EVENTO.WebLocalWhere = "A.INCLUINOTIPOANEXO = 'Q' AND A.ULTIMONIVEL = 'S' AND A.INATIVO <> 'S' AND EXISTS(SELECT 1 FROM SAM_TGE_TABELATISS T WHERE T.EVENTO = A.HANDLE AND T.TABELATISS = @CAMPO(CODIGOTABELA))"

  Set vSQL = Nothing
End Sub

Public Function PegarHandleVersaoTISS(Optional piHandleAutorizacao As Long = 0) As Integer
    Dim viVersaoTISS
    Dim sql As Object
    Set sql = NewQuery

	If piHandleAutorizacao <> 0 Then
      sql.Active = False
	  sql.Clear
      sql.Add("SELECT A.VERSAOTISS VERSAOTISS")
      sql.Add("FROM SAM_AUTORIZ A ")
      sql.Add("WHERE A.HANDLE = :HANDLE")
      sql.ParamByName("HANDLE").Value = piHandleAutorizacao
      sql.Active = True

	  If sql.FieldByName("VERSAOTISS").AsInteger = 0 Then
        sql.Active = False
	    sql.Clear
        sql.Add("SELECT MAX(A.HANDLE) VERSAOTISS ")
        sql.Add("FROM TIS_VERSAO A ")
        sql.Add("WHERE A.ATIVODESKTOP = 'S' ")
        sql.Active = True
	  End If
	Else
      sql.Active = False
      sql.Clear
      sql.Add("SELECT MAX(A.HANDLE) VERSAOTISS ")
      sql.Add("FROM TIS_VERSAO A ")
      sql.Add("WHERE A.ATIVODESKTOP = 'S' ")
      sql.Active = True
	End If

    viVersaoTISS = sql.FieldByName("VERSAOTISS").AsInteger
	Set sql = Nothing

	PegarHandleVersaoTISS = viVersaoTISS
End Function

