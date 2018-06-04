'HASH: 45DB27505201DAD018CDEE220D556A52
Sub Main() 
  Dim Executar As Object 
  Dim Q        As Object
  Set Q = NewQuery
  Q.Add("SELECT USUARIOPADRAO, HOSTPADRAO FROM AEX_PARAMETROSGERAIS")
  Q.Active=True 
  Set Executar = CreateBennerObject("BSAte006.Rotinas")
  Executar.ProcessarAgendamento(CurrentSystem,"",356, _
                                Q.FieldByName("USUARIOPADRAO").AsInteger, Q.FieldByName("HOSTPADRAO").AsString)
  Set Executar = Nothing
  Set Q = Nothing
End Sub
