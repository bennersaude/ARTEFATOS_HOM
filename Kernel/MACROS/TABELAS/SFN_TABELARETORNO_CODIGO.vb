'HASH: C157D7C8CA72DF9E3210E8D0B2997F7D
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("DEBGERARDOCUMENTO").AsString = "S" Then
    If CurrentQuery.FieldByName("DEBTIPODOCUMENTO").AsString = "" Then
      bsShowMessage("Campo Tipo Documento Obrigatório!", "I")
      DEBTIPODOCUMENTO.SetFocus
    End If
  End If
  If CurrentQuery.FieldByName("CREGERARDOCUMENTO").AsString = "S" Then
    If CurrentQuery.FieldByName("CRETIPODOCUMENTO").AsString = "" Then
      bsShowMessage("Campo Tipo Documento Obrigatório!", "I")
      CRETIPODOCUMENTO.SetFocus
    End If
  End If


' Coelho SMS: 77998 retirado o trecho abaixo pois o campo EFETUARBAIXA NÃO SERÁ MAIS UTILIZADO
'  If (CurrentQuery.FieldByName("TIPORETORNO").AsString = "C" Or CurrentQuery.FieldByName("TIPORETORNO").AsString = "S") And _
'      (CurrentQuery.FieldByName("EFETUARBAIXA").AsString = "S") Then
'    MsgBox ("Tipo de Retorno incompatível com Efetua Baixa!")
'    CurrentQuery.FieldByName("EFETUARBAIXA").AsString = "N"
'  End If


' Coelho SMS: 77998 ALTERADO o trecho abaixo pois o campo EFETUARBAIXA NÃO SERÁ MAIS UTILIZADO
 'LOPES
 ' If CurrentQuery.FieldByName("TABREENVIARNOVOVENCIMENTO").AsInteger = 2 Then
 '   If (CurrentQuery.FieldByName("TIPORETORNO").AsString <> "S") Or (CurrentQuery.FieldByName("EFETUARBAIXA").AsString <> "N") Then
 '     MsgBox("Para reenviar com novo vencimento o tipo de retorno deve estar configurado com a opção 'Erro sem nova remessa' e o parâmetro 'Efetuar baixa' deve estar desmarcado !", vbCritical, "Benner Saúde")
 '     CanContinue = False
 '   End If
 ' End If

 'Coelho SMS: 77998 alteração do trecho acima para o trecho abaixo pois o campo EFETUARBAIXA NÃO SERÁ MAIS UTILIZADO
  If CurrentQuery.FieldByName("TABREENVIARNOVOVENCIMENTO").AsInteger = 2 Then
    If (CurrentQuery.FieldByName("TIPORETORNO").AsString <> "S") Then
      bsShowMessage("Para reenviar com novo vencimento o tipo de retorno deve estar configurado com a opção 'Erro sem nova remessa'!", "E")
      CanContinue = False
    End If
  End If


End Sub

