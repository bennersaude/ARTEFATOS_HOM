'HASH: 6F82E5299E10B21E34FF841E2B5DC9F4
 
'Macro: SAM_MOTIVOMOVTOACERTO

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  'Soares - SMS: 62751 - 10/08/2006 - Início
  Dim qVerificaMesmoCodigo As Object

  Set qVerificaMesmoCodigo = NewQuery
  qVerificaMesmoCodigo.Add("SELECT CODIGO               ")
  qVerificaMesmoCodigo.Add("  FROM SAM_MOTIVOMOVTOACERTO")
  qVerificaMesmoCodigo.Add(" WHERE CODIGO =  :CODIGO    ")
  qVerificaMesmoCodigo.Add("   AND HANDLE <> :HANDLE    ") 'Se editar e salvar não exibe a mensagem
  qVerificaMesmoCodigo.ParamByName("CODIGO").AsInteger = CurrentQuery.FieldByName("CODIGO").AsInteger
  qVerificaMesmoCodigo.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qVerificaMesmoCodigo.Active = True

  ' Verifica se já há alguem cadastrado com esse codigo, se tiver não aceita inserir novamente
  If (Not qVerificaMesmoCodigo.FieldByName("CODIGO").IsNull) Then
    MsgBox("Já existe um Motivo de Movimento de Acerto cadastrado com esse código")
    CanContinue = False
    CODIGO.SetFocus
    Set qVerificaMesmoCodigo = Nothing
    Exit Sub
  End If

  Set qVerificaMesmoCodigo = Nothing
'Soares - SMS: 62751 - 10/08/2006 - Fim
End Sub
