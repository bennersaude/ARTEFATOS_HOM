﻿'HASH: 74F747354741979FEC5232D0E6B54B46

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)


  If WebMode Then
  	FRANQUIAGRP.WebLocalWhere = "FRANQUIA IN (SELECT FGRP.FRANQUIA FROM SAM_FRANQUIAGRP FGRP JOIN SAM_PLANO_FRANQUIA PF ON (PF.FRANQUIA = FGRP.FRANQUIA) WHERE PF.HANDLE = @CAMPO(PLANOFRANQUIA))"
  ElseIf VisibleMode Then
  	FRANQUIAGRP.LocalWhere = "FRANQUIA IN (SELECT FGRP.FRANQUIA FROM SAM_FRANQUIAGRP FGRP JOIN SAM_PLANO_FRANQUIA PF ON (PF.FRANQUIA = FGRP.FRANQUIA) WHERE PF.HANDLE = @PLANOFRANQUIA)"
  End If

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If WebMode Then
  	FRANQUIAGRP.WebLocalWhere = "FRANQUIA IN (SELECT FGRP.FRANQUIA FROM SAM_FRANQUIAGRP FGRP JOIN SAM_PLANO_FRANQUIA PF ON (PF.FRANQUIA = FGRP.FRANQUIA) WHERE PF.HANDLE = @CAMPO(PLANOFRANQUIA))"
  ElseIf VisibleMode Then
    FRANQUIAGRP.LocalWhere = "FRANQUIA IN (SELECT FGRP.FRANQUIA FROM SAM_FRANQUIAGRP FGRP JOIN SAM_PLANO_FRANQUIA PF ON (PF.FRANQUIA = FGRP.FRANQUIA) WHERE PF.HANDLE = @PLANOFRANQUIA)"
  End If
End Sub
