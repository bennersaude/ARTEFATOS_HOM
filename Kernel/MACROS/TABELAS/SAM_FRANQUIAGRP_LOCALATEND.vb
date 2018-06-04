'HASH: 5DE63B8B60E7AD5B091E9B2B937FE500
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Add("SELECT *                                ")
Consulta.Add("  FROM SAM_FRANQUIAGRP_LOCALATEND       ")
Consulta.Add(" WHERE FRANQUIAGRP = :FRANQUIAGRP       ")
Consulta.Add("   AND LOCALATENDIMENTO = :LOCAL        ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE                ")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If
Consulta.ParamByName("FRANQUIAGRP").AsInteger = CurrentQuery.FieldByName("FRANQUIAGRP").AsInteger
Consulta.ParamByName("LOCAL").AsInteger = CurrentQuery.FieldByName("LOCALATENDIMENTO").AsInteger
Consulta.Active = True

If Not Consulta.FieldByName("HANDLE").IsNull Then
  bsShowMessage("Local de Atendimento já cadastrado!", "E")
  CanContinue = False
  Exit Sub
End If
End Sub
