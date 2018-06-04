'HASH: 94CA04AC794B2FD405824CA77B7FAA57
'#Uses "*bsShowMessage"

Option Explicit
Public Sub CLASSEGERENCIAL_OnPopup(ShowPopup As Boolean)


  Dim interface As Object

  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String


  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_CLASSEGERENCIAL.ESTRUTURA|SFN_CLASSEGERENCIAL.DESCRICAO|SFN_CLASSEGERENCIAL.CODIGOREDUZIDO|SFN_CLASSEGERENCIAL.NATUREZA|SFN_CLASSEGERENCIAL.HISTORICO"

  vCriterio = "HANDLE>0"

  vCampos = "Estrutura|Descrição|Código|D/C|Historico"

  vHandle = interface.Exec(CurrentSystem, "SFN_CLASSEGERENCIAL", vColunas, 1, vCampos, vCriterio, "Classe Gerencial", False, CLASSEGERENCIAL.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CLASSEGERENCIAL").Value = vHandle
  End If
  Set interface = Nothing

End Sub


Public Sub TABLE_AfterScroll()
	If WebMode Then
		CLASSEGERENCIAL.WebLocalWhere = "HANDLE>0"
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
 Dim SQL As Object

  Set SQL = NewQuery

  SQL.Add("SELECT C.CODIGO, C.DESCRICAO, CG.HANDLE")
  SQL.Add(" FROM SFN_CODIGOFOLHA C, SFN_CODIGOFOLHA_CLASSEGERENCIA CG")
  SQL.Add(" WHERE C.HANDLE = CG.CODIGOFOLHA")
  SQL.Add(" AND CG.CLASSEGERENCIAL = :CLASSE")
  'SQL.Add(" AND CG.OPERACAO = :OPERACAO")
  SQL.Add(" AND CG.CODIGOFOLHA <> :CODIGO")

  SQL.ParamByName("CLASSE").Value = CurrentQuery.FieldByName("CLASSEGERENCIAL").AsInteger
  'SQL.ParamByName("OPERACAO").Value = CurrentQuery.FieldByName("OPERACAO").AsInteger
  SQL.ParamByName("CODIGO").Value = CurrentQuery.FieldByName("CODIGOFOLHA").AsInteger
  SQL.Active = True

  If Not SQL.EOF Then
    bsShowMessage("Já existe um Código folha cadastrado com essa classe gerencial e operação ("+ SQL.FieldByName("CODIGO").AsString + ")", "E")
    CanContinue = False
  End If

  Set SQL = Nothing
  
End Sub
