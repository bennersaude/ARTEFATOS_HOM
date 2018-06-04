'HASH: DD2D690B5F8A94F2308E96D96FAD2B98
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)

Dim q As Object
Set q = NewQuery

Dim qAux As Object
Set qAux = NewQuery
Dim qTemp As Object
Set qTemp = NewQuery



q.Clear
q.Add("SELECT SUM(M.REMESSAENTRADA) REMESSAENTRADA, SUM(M.REMESSABAIXA) REMESSABAIXA, SUM(M.REMESSACANCELA) REMESSACANCELA, SUM(M.REMESSAALTERACAOVENCIMENTO) REMESSAALTERACAOVENCIMENTO")
q.Add("  FROM SFN_MODELO M, SFN_TIPODOCUMENTO_MODELO MD")
q.Add(" WHERE M.HANDLE = MD.MODELODOCUMENTO")
q.Add("   AND MD.TIPODOCUMENTO = :TIPO")
q.Add("   AND MD.HANDLE        <> :HANDLE")
q.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
q.ParamByName("TIPO").AsInteger   = CurrentQuery.FieldByName("TIPODOCUMENTO").AsInteger
q.Active = True

qAux.Clear
qAux.Add("SELECT M.REMESSAENTRADA, M.REMESSABAIXA, M.REMESSACANCELA, M.REMESSAALTERACAOVENCIMENTO")
qAux.Add("  FROM SFN_MODELO M")
qAux.Add(" WHERE M.HANDLE = :HANDLE")
qAux.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("MODELODOCUMENTO").AsInteger
qAux.Active = True

If (q.FieldByName("REMESSAENTRADA").AsInteger > 0) And (qAux.FieldByName("REMESSAENTRADA").AsInteger > 0) Then
   qTemp.Clear
   qTemp.Add("SELECT M.DESCRICAO")
   qTemp.Add("  FROM SFN_MODELO M")
   qTemp.Add("  JOIN SFN_TIPODOCUMENTO_MODELO TM ON (TM.MODELODOCUMENTO = M.HANDLE)")
   qTemp.Add(" WHERE TM.TIPODOCUMENTO = :TIPODOCUMENTO AND M.REMESSAENTRADA > 0")
   qTemp.ParamByName("TIPODOCUMENTO").AsInteger = CurrentQuery.FieldByName("TIPODOCUMENTO").AsInteger
   qTemp.Active = True
   bsShowMessage("Modelo "+qTemp.FieldByName("DESCRICAO").AsString+" com código de envio de Remessa já informado no tipo de documento!", "E")
   CanContinue = False
   Exit Sub
End If

If (q.FieldByName("REMESSABAIXA").AsInteger > 0) And (qAux.FieldByName("REMESSABAIXA").AsInteger > 0) Then
   qTemp.Close
   qTemp.SQL.Clear
   qTemp.SQL.Add("SELECT M.DESCRICAO")
   qTemp.SQL.Add("  FROM SFN_MODELO M")
   qTemp.SQL.Add("  JOIN SFN_TIPODOCUMENTO_MODELO TM ON (TM.MODELODOCUMENTO = M.HANDLE)")
   qTemp.SQL.Add(" WHERE TM.TIPODOCUMENTO = :TIPODOCUMENTO AND M.REMESSABAIXA > 0")
   qTemp.ParamByName("TIPODOCUMENTO").AsInteger = CurrentQuery.FieldByName("TIPODOCUMENTO").AsInteger
   qTemp.Active = True
   bsShowMessage("Modelo "+qTemp.FieldByName("DESCRICAO").AsString+" com código de envio de Baixa já informado no tipo de documento!", "E")
   CanContinue = False
   Exit Sub
End If

If (q.FieldByName("REMESSACANCELA").AsInteger > 0) And (qAux.FieldByName("REMESSACANCELA").AsInteger > 0) Then
   qTemp.Close
   qTemp.SQL.Clear
   qTemp.SQL.Add("SELECT M.DESCRICAO")
   qTemp.SQL.Add("  FROM SFN_MODELO M")
   qTemp.SQL.Add("  JOIN SFN_TIPODOCUMENTO_MODELO TM ON (TM.MODELODOCUMENTO = M.HANDLE)")
   qTemp.SQL.Add(" WHERE TM.TIPODOCUMENTO = :TIPODOCUMENTO AND M.REMESSACANCELA > 0")
   qTemp.ParamByName("TIPODOCUMENTO").AsInteger = CurrentQuery.FieldByName("TIPODOCUMENTO").AsInteger
   qTemp.Active = True
   bsShowMessage("Modelo "+qTemp.FieldByName("DESCRICAO").AsString+" com código de envio de Cancelamento já informado no tipo de documento!", "E")
   CanContinue = False
   Exit Sub
End If

If (q.FieldByName("REMESSAALTERACAOVENCIMENTO").AsInteger > 0) And (qAux.FieldByName("REMESSAALTERACAOVENCIMENTO").AsInteger > 0) Then
   qTemp.Close
   qTemp.SQL.Clear
   qTemp.SQL.Add("SELECT M.DESCRICAO")
   qTemp.SQL.Add("  FROM SFN_MODELO M")
   qTemp.SQL.Add("  JOIN SFN_TIPODOCUMENTO_MODELO TM ON (TM.MODELODOCUMENTO = M.HANDLE)")
   qTemp.SQL.Add(" WHERE TM.TIPODOCUMENTO = :TIPODOCUMENTO AND M.REMESSAALTERACAOVENCIMENTO > 0")
   qTemp.ParamByName("TIPODOCUMENTO").AsInteger = CurrentQuery.FieldByName("TIPODOCUMENTO").AsInteger
   qTemp.Active = True
   bsShowMessage("Modelo "+qTemp.FieldByName("DESCRICAO").AsString+" com código de envio de Remessa de Alteração de Vencimento já informado no tipo de documento!", "E")
   CanContinue = False
   Exit Sub
End If

End Sub
