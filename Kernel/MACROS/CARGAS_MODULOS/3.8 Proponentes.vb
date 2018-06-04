'HASH: 11036DAB7BB737D25F191C34522E8C4D
 

Public Sub CONSULTAESPECIALIDAD_OnClick()
  Dim OLESamConsulta As Object
  Set OLESamObject =CreateBennerObject("SamConsulta.Consulta")
  OLESamObject.ProponenteEspecialidade(CurrentSystem)
  Set OLESamObject =Nothing

End Sub
