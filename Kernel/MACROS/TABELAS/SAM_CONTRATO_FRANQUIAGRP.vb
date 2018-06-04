'HASH: 6E3C0A893AF0F2066B8D7A5592D394F0
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If WebMode Then
		FRANQUIAGRP.WebLocalWhere = "FRANQUIA IN (SELECT F.FRANQUIA FROM SAM_FRANQUIAGRP F JOIN SAM_CONTRATO_FRANQUIA CF ON (F.FRANQUIA = CF.FRANQUIA) WHERE CF.HANDLE = @CAMPO(CONTRATOFRANQUIA))"
	ElseIf VisibleMode Then
		FRANQUIAGRP.LocalWhere = "FRANQUIA IN (SELECT F.FRANQUIA FROM SAM_FRANQUIAGRP F JOIN SAM_CONTRATO_FRANQUIA CF ON (F.FRANQUIA = CF.FRANQUIA) WHERE CF.HANDLE = @CONTRATOFRANQUIA)"
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	If WebMode Then
		FRANQUIAGRP.WebLocalWhere = "FRANQUIA IN (SELECT F.FRANQUIA FROM SAM_FRANQUIAGRP F JOIN SAM_CONTRATO_FRANQUIA CF ON (F.FRANQUIA = CF.FRANQUIA) WHERE CF.HANDLE = @CAMPO(CONTRATOFRANQUIA))"
	ElseIf VisibleMode Then
		FRANQUIAGRP.LocalWhere = "FRANQUIA IN (SELECT F.FRANQUIA FROM SAM_FRANQUIAGRP F JOIN SAM_CONTRATO_FRANQUIA CF ON (F.FRANQUIA = CF.FRANQUIA) WHERE CF.HANDLE = @CONTRATOFRANQUIA)"
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qSel As Object
  Set qSel = NewQuery

  qSel.Active = False
  qSel.Add("SELECT COUNT(HANDLE) QTDE ")
  qSel.Add("  FROM SAM_CONTRATO_FRANQUIAGRP")
  qSel.Add(" WHERE FRANQUIAGRP = :FRANQUIAGRP")
  qSel.Add("       AND HANDLE <> :HANDLE")
  qSel.Add("       AND CONTRATOFRANQUIA = :CONTRATOFRANQUIA ")
  qSel.ParamByName("FRANQUIAGRP").AsInteger = CurrentQuery.FieldByName("FRANQUIAGRP").AsInteger
  qSel.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qSel.ParamByName("CONTRATOFRANQUIA").AsInteger = CurrentQuery.FieldByName("CONTRATOFRANQUIA").AsInteger
  qSel.Active = True

  If qSel.FieldByName("QTDE").AsInteger > 0 Then
    bsShowMessage("Grupo de franquia já cadastrado na franquia.", "E")
    Set qSel = Nothing
    CanContinue = False
    Exit Sub
  End If

  Set qSel = Nothing
End Sub

