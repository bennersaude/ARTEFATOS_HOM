'HASH: 12BD3791ADFECFAE030484FBD9192AD1
 

Public Sub BOTAOGERAREGRAS_OnClick()
Dim Interface As Object
  Set Interface =CreateBennerObject("DupEvento.DupEventos")
  Interface.RegraExcecao(CurrentSystem,"SAM_PRESTADOR_REGRA",RecordHandleOfTable("SAM_PRESTADOR"),"R")
  Set Interface =Nothing
  RefreshNodesWithTable "SAM_PRESTADOR_REGRA"
End Sub



Public Sub CADASTRARREGEXC_OnClick()

  Dim Interface As Object
  Dim SQL As Object


  Set Interface =CreateBennerObject("SamProcPrestador.ProcessoPrestador")

  Set SQL =NewQuery

  SQL.Add("SELECT EDITAREGRAEXCECAO FROM SAM_PARAMETROSPRESTADOR")
  SQL.Active =True

  If SQL.FieldByName("EDITAREGRAEXCECAO").Value ="S" Then
    Interface.RegraExcPrestFiliados(CurrentSystem,RecordHandleOfTable("SAM_PRESTADOR"))
  Else
    MsgBox "A opção 'Edita regra exceção' está desmarcada  (Carga: Adm/Parâmetros Gerais/Prestadores)  !"
  End If

  Set Interface =Nothing

End Sub
