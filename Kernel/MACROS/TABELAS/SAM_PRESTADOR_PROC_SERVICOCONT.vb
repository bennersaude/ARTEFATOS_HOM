'HASH: 1DC94D449FAC6A187706559DD6E32C9F
Option Explicit

'#Uses "*bsShowMessage"

Public Sub SERVICOCONTRATADO_OnPopup(ShowPopup As Boolean)
	SERVICOCONTRATADO.LocalWhere = Condicao
End Sub

Public Sub SERVICOCONTRATADOEXCLUIR_OnPopup(ShowPopup As Boolean)
	SERVICOCONTRATADOEXCLUIR.LocalWhere = Condicao
End Sub


Public Sub TABLE_AfterScroll()
  PRESTADOR.Visible = False
End Sub

Public Function Condicao As String

	CurrentQuery.UpdateRecord

	Dim qSql      As String
	Dim vCondicao As String

	If CurrentQuery.FieldByName("TABOPERACAO").AsInteger = 1 Then
		vCondicao = "NOT EXISTS "
	Else
		vCondicao = "EXISTS "
	End If


	qSql = qSql &  vCondicao & "(SELECT S.HANDLE"
	qSql = qSql &  "               FROM SAM_PRESTADOR_SERVICOSCONTRAT S"
	qSql = qSql &  "              WHERE S.PRESTADOR = " & CurrentQuery.FieldByName("PRESTADOR").AsString
	qSql = qSql &  "                AND S.SERVICOCONTRATADO = A.HANDLE)"

	Condicao = qSql

End Function

Public Sub TABLE_NewRecord()
  CurrentQuery.FieldByName("RESPONSAVELANALISE").Value = CurrentUser

  Dim buscaPrestador As Object
  Set buscaPrestador = NewQuery


  buscaPrestador.Add("SELECT PRESTADOR           ")
  buscaPrestador.Add("  FROM SAM_PRESTADOR_PROC  ")
  buscaPrestador.Add(" WHERE HANDLE = :PROCESSO  ")

  buscaPrestador.ParamByName("PROCESSO").AsInteger = RecordHandleOfTable("SAM_PRESTADOR_PROC")
  buscaPrestador.Active = True

  CurrentQuery.FieldByName("PRESTADOR").AsInteger = buscaPrestador.FieldByName("PRESTADOR").AsInteger

  Set buscaPrestador = Nothing
End Sub
