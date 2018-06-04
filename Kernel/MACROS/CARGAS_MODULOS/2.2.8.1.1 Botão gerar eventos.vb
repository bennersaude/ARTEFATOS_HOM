'HASH: E952D30A14FBFBB980750D4BB60E16AC
 

Public Sub BOTAOGERAREVENTOS_OnClick()
  
  Dim Duplica As Object
  Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
  Duplica.Duplicar(CurrentSystem,"SAM_FRANQUIAGRP_EVENTO","FRANQUIAGRP",RecordHandleOfTable("SAM_FRANQUIAGRP"),"Duplicando eventos para franquia")
  Set Duplica =Nothing
  RefreshNodesWithTable "SAM_FRANQUIAGRP_EVENTO"

End Sub
