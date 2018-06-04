'HASH: 723A4854FC18A58A4DDAB555C6816BB5
'Macro: SAM_ROTAVISOSUSPENSAO_CONTRATO


Public Sub TABLE_AfterInsert()
  Dim ROTINA As Object
  Set ROTINA = NewQuery

  ROTINA.Clear
  ROTINA.Add("Select PROCESSADO")
  ROTINA.Add("  FROM SAM_ROTAVISOSUSPENSAO")
  ROTINA.Add(" WHERE HANDLE = :HANDLE")
  ROTINA.ParamByName("HANDLE").Value = RecordHandleOfTable(SAM_ROTAVISOSUSPENSAO)
  ROTINA.Active = True

  If ROTINA.FieldByName("PROCESSADO").AsString = "S" Then
    CANCONTINUE = False
  End If

  Set ROTINA = Nothing
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

End Sub

