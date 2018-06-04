'HASH: 6AA67DA0564A924A757CBD36E2EA3CFC
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_AfterScroll()
  USUARIO.ReadOnly = True
  DATAREGISTRO.ReadOnly = True
  OBSERVACAOINCLUIDAPELOUSUARIO.ReadOnly = True
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  CanContinue = PermiteOperacoes
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  CanContinue = PermiteOperacoes
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  CanContinue = PermiteOperacoes
End Sub

Function PermiteOperacoes As Boolean
  PermiteOperacoes = True
  Dim sql As Object
  Dim mensagem As String

  Set sql=NewQuery

  sql.Add("SELECT SITUACAO FROM SAM_AUTORIZ WHERE HANDLE = :AUTORIZACAO")
  sql.ParamByName("AUTORIZACAO").AsInteger = RecordHandleOfTable("SAM_AUTORIZ")
  sql.Active = True

  If sql.FieldByName("SITUACAO").AsString = "F" Or sql.FieldByName("SITUACAO").AsString = "C" Then
    mensagem = "Não é possível realizar esta operação pois a autorização está "
    If sql.FieldByName("SITUACAO").AsString = "F" Then
      mensagem = mensagem + " fechada."
    Else
      mensagem = mensagem + " cancelada."
    End If
    BsShowMessage(mensagem, "E")
    PermiteOperacoes = False
  End If
End Function

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ") > 0 Then
    CurrentQuery.FieldByName("PROTOCOLOTRANSACAO").AsInteger = RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ")
  End If

  CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser
  CurrentQuery.FieldByName("DATAREGISTRO").AsDateTime = ServerNow
  CurrentQuery.FieldByName("OBSERVACAOINCLUIDAPELOUSUARIO").AsString = "S"
End Sub

Public Sub TABLE_NewRecord()
  CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser
  CurrentQuery.FieldByName("DATAREGISTRO").AsDateTime = ServerNow
End Sub
