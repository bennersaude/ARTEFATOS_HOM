'HASH: D0538E798546171BA721FB071D446C1E
 '#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Add("SELECT *                                ")
Consulta.Add("  FROM SAM_FRANQUIAGRP_CONDATEND        ")
Consulta.Add(" WHERE FRANQUIAGRP = :FRANQUIAGRP       ")
Consulta.Add("   AND CONDICAOATENDIMENTO = :CONDICAO  ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE                ")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If
Consulta.ParamByName("FRANQUIAGRP").AsInteger = CurrentQuery.FieldByName("FRANQUIAGRP").AsInteger
Consulta.ParamByName("CONDICAO").AsInteger = CurrentQuery.FieldByName("CONDICAOATENDIMENTO").AsInteger
Consulta.Active = True

If Not Consulta.FieldByName("HANDLE").IsNull Then
  bsShowMessage("Condição de atendimento já cadastrado!", "E")
  CanContinue = False
  Exit Sub
End If
End Sub
