'HASH: B658F101FD85D686E3228AECE8732A9E
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Add("SELECT *                                ")
Consulta.Add("  FROM SAM_FRANQUIAGRP_FINATENDIMENTO   ")
Consulta.Add(" WHERE FRANQUIAGRP = :FRANQUIAGRP       ")
Consulta.Add("   AND FINALIDADEATENDIMENTO = :FINAL   ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE                ")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If
Consulta.ParamByName("FRANQUIAGRP").AsInteger = CurrentQuery.FieldByName("FRANQUIAGRP").AsInteger
Consulta.ParamByName("FINAL").AsInteger = CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").AsInteger
Consulta.Active = True

If Not Consulta.FieldByName("HANDLE").IsNull Then
  bsShowMessage("Finalidade já cadastrada!", "E")
  CanContinue = False
  Exit Sub
End If
End Sub
