'HASH: A22174B57AE0FA348A75874A078C57AE
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Add("SELECT *                                ")
Consulta.Add("  FROM SAM_FRANQUIAGRP_REGATENDIMENTO   ")
Consulta.Add(" WHERE FRANQUIAGRP = :FRANQUIAGRP       ")
Consulta.Add("   AND REGIMEATENDIMENTO = :REGIME      ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE                ")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If
Consulta.ParamByName("FRANQUIAGRP").AsInteger = CurrentQuery.FieldByName("FRANQUIAGRP").AsInteger
Consulta.ParamByName("REGIME").AsInteger = CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger
Consulta.Active = True

If Not Consulta.FieldByName("HANDLE").IsNull Then
  bsShowMessage("Regime de Atendimento já cadastrado!", "E")
  CanContinue = False
  Exit Sub
End If
End Sub
