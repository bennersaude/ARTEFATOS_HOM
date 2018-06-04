'HASH: 0858149FBA7B266815484EC08559898E
'#Uses "*bsShowMessage"

Option Explicit

Dim vHandleAutorizacao As Long
Dim vCobrarValorExcedente As String

Public Sub TABLE_NewRecord()

  Dim qAutoriz As BPesquisa

  vHandleAutorizacao = 0

  If SessionVar("HANDLEAUTORIZACAO") <> "" Then
    vHandleAutorizacao = CLng(SessionVar("HANDLEAUTORIZACAO"))

    Set qAutoriz = NewQuery

    qAutoriz.Clear
    qAutoriz.Active = False
    qAutoriz.Add("SELECT COBRARVALOREXCEDENTE ")
    qAutoriz.Add("  FROM SAM_AUTORIZ          ")
    qAutoriz.Add(" WHERE HANDLE = :HANDLE     ")

    qAutoriz.ParamByName("HANDLE").AsInteger = vHandleAutorizacao
    qAutoriz.Active = True

    vCobrarValorExcedente = qAutoriz.FieldByName("COBRARVALOREXCEDENTE").AsString

    CurrentQuery.FieldByName("COBRARVALOREXCEDENTE").AsString = vCobrarValorExcedente

    Set qAutoriz = Nothing
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  On Error GoTo erro
    Dim callEntity As CSEntityCall
    Set callEntity = BusinessEntity.CreateCall("Benner.Saude.Entidades.Atendimento.TvValorExcedente, Benner.Saude.Entidades", "AlterarIndicadorCobrancaValorExcedente")
    callEntity.AddParameter(pdtAutomatic, vHandleAutorizacao)
    callEntity.AddParameter(pdtAutomatic, vCobrarValorExcedente)
    callEntity.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("COBRARVALOREXCEDENTE").AsString)
    callEntity.Execute
    Set callEntity = Nothing

    Exit Sub
  erro:
    Set callEntity = Nothing
    bsShowMessage("Ocorreu o seguinte erro ao alterar o campo: " + Str(Error), "E")
    CanContinue = False
    Exit Sub
End Sub
