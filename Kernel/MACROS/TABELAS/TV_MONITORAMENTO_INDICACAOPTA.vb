'HASH: 68B30F15A95955D01A28D0F9C0B5C658
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_AfterPost()

  'A entidade desta tabela virtual já está implementada, porém o processo está sendo feito pela macro, pois pelo WES2006 a entidaade não recebe uma sessionvar

  Dim verificarRepitido As BPesquisa
  Set verificarRepitido = NewQuery
  verificarRepitido.Add("SELECT HANDLE                       ")
  verificarRepitido.Add("  FROM ANS_TISMONITORAMENTO         ")
  verificarRepitido.Add(" WHERE HANDLE  <> :HANDLE           ")
  verificarRepitido.Add("   AND PROTOCOLOPTA = :PROTOCOLOPTA ")

  verificarRepitido.ParamByName("HANDLE").AsString = SessionVar("HANDLE_ROTMONITORAMENTOTISS")
  verificarRepitido.ParamByName("PROTOCOLOPTA").AsString = CurrentQuery.FieldByName("PROTOCOLOPTA").AsString
  verificarRepitido.Active = True

  If Not verificarRepitido.EOF And (CurrentQuery.FieldByName("PROTOCOLOPTA").AsString <> "") Then
    bsShowMessage("O protocolo PTA informado já existe em outra rotina!","I")
    Exit Sub
  End If

  Set verificarRepitido = Nothing

  Dim bs As CSBusinessComponent
  Dim handleProcesso As Long

  Set bs = BusinessComponent.CreateInstance("Benner.Saude.ANS.Processos.Monitoramento.Rotina.IndicarProtocoloPta, Benner.Saude.ANS.Processos")
  bs.AddParameter(pdtString, SessionVar("HANDLE_ROTMONITORAMENTOTISS"))
  bs.AddParameter(pdtString, CurrentQuery.FieldByName("PROTOCOLOPTA").AsString)

  bs.Execute("AlterarProtocoloPta")

  Set bs = Nothing

End Sub
