'HASH: B85DBDDFCE6F128EA2B4AE6DBC4A60C6
'#Uses "*RegistrarLogAlteracao"
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_AfterPost()
  RegistrarLogAlteracao "SAM_REGISTROMS", CurrentQuery.FieldByName("HANDLE").AsInteger, "TABLE_AfterPost"
End Sub

Public Sub TABLE_AfterScroll()
  'Daniela - SMS 12220 - Convênio no registro da ANS
  If Not CurrentQuery.FieldByName("CONVENIO").IsNull Then
    CONVENIO.ReadOnly = True
  Else
    CONVENIO.ReadOnly = False
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("NOVAREGULAMENTACAO").AsString = "S" And CurrentQuery.FieldByName("ENVIADOSCPA").AsString = "S" Then
    bsShowMessage("Não é possível marcar o campo Enviado SCPA se o campo Nova Regulamentação estiver marcado","E")
    CanContinue = False
  End If

  If Not CurrentQuery.FieldByName("DATAVENCIMENTO").IsNull Then
    If CurrentQuery.FieldByName("DATAREGISTRO").AsDateTime > CurrentQuery.FieldByName("DATAVENCIMENTO").AsDateTime Then
      bsShowMessage("A Data de vencimento não pode ser menor que a data de Registro!","E")
      CanContinue = False
    End If
  End If

End Sub
