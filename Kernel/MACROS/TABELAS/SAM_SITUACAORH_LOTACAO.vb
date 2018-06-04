'HASH: FBE835277288BF72486A03E05987D5D7
 


Public Sub LOTACAODESTINO_OnPopup(ShowPopup As Boolean)
  Dim vSelect As String
  vSelect = "HANDLE In (SELECT A.HANDLE" + Chr(13) + Chr(10)
  vSelect = vSelect + "FROM SAM_CONTRATO_LOTACAO A, "+ Chr(13) + Chr(10)
  vSelect = vSelect + "     SAM_SITUACAORH       B  "+ Chr(13) + Chr(10)
  vSelect = vSelect + "WHERE A.CONTRATO = B.CONTRATOMIGRACAO"+ Chr(13) + Chr(10)
  VSelect = vSelect + "AND B.HANDLE = " + CurrentQuery.FieldByName("SITUACAORH").AsString + ")"

  LOTACAODESTINO.LocalWhere = vSelect


End Sub

Public Sub LOTACAOORIGEM_OnPopup(ShowPopup As Boolean)
  Dim vSelect As String
  vSelect = "HANDLE In (SELECT A.HANDLE" + Chr(13) + Chr(10)
  vSelect = vSelect + "FROM SAM_CONTRATO_LOTACAO A, "+ Chr(13) + Chr(10)
  vSelect = vSelect + "     SAM_SITUACAORH       B  "+ Chr(13) + Chr(10)
  vSelect = vSelect + "WHERE A.CONTRATO = B.CONTRATO"+ Chr(13) + Chr(10)
  vSelect = vSelect + "AND B.HANDLE = " + CurrentQuery.FieldByName("SITUACAORH").AsString + ")"


  LOTACAOORIGEM.LocalWhere = vSelect


End Sub
