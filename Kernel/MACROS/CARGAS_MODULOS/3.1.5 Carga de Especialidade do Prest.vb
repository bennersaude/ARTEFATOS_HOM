'HASH: 188E9F05956E32107306467638B6CF0F
 

Public Sub INSERIRESPEC_OnClick()

  Dim Interface As Object
  Dim SQL As Object


  Set Interface =CreateBennerObject("SamProcPrestador.ProcessoPrestador")

  Set SQL =NewQuery

  SQL.Add("SELECT EDITAESPECIALIDADE FROM SAM_PARAMETROSPRESTADOR")
  SQL.Active =True

  If SQL.FieldByName("EDITAESPECIALIDADE").Value ="S" Then
    Interface.EspecialidadePrestFiliados(CurrentSystem,RecordHandleOfTable("SAM_PRESTADOR"))
  Else
    MsgBox "A opção 'Edita especialidade' está desmarcada  (Carga: Adm/Parâmetros Gerais/Prestadores)  !"
  End If

  Set Interface =Nothing

End Sub
