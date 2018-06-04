'HASH: D4FC7B3F4B13F7EB8E06B1DE28F73D64


Public Sub BOTAOGERACALENDARIOPGTO_OnClick()
  '-------------------------------
  ' Rotina criada em 15/09/2003
  ' Programador: Douglas D.Lara
  ' SMS nº 15313
  ' Pagamento de guia
  '-------------------------------
  ' Rotina que faz com que seja instanciada a dll SamCalendarioPagamento,
  ' que abre um Form para que sejam dados parâmetros para a geração
  ' automâtica de datas de pagamento e sua associação com calendários de
  ' recebimento de PEGs.
  ' Caso o registro atual(SAM_CALENDGERAL)esteja em edição,ele dá mensagem
  ' de erro,e só permite a execução caso o registro não esteja em edição,pois
  ' é necessário o handle do registro em SAM_CALENDGERAL para que sejam gerados
  ' registros na tabela SAM_CALENDGERAL_RECEBIMENTO,que se relaciona com a tabela
  ' SAM_CALENDGERAL através da coluna CALENDGERAL e com a tabela SAM_PAGAMENTO
  ' através da coluna PAGAMENTO.
  '--------------------------------

  Dim CriaCalendRecebimento As Object
  Dim CalendarioAtual As Integer

  If CurrentQuery.State = 1 Then 'Query não em edição ->>Instancia dll e executa
    CalendarioAtual = CurrentQuery.FieldByName("CODIGO").AsInteger
    Set CalendarioPgto = CreateBennerObject("SamCalendarioPgto.Rotinas")
    CalendarioPgto.CriaCalendPagamento(CurrentSystem, CalendarioAtual)
    Set CalendarioPgto = Nothing

  Else 'Query em edição ->>Erro!
    MsgBox("O Registro não pode estar em edição!")
  End If

End Sub

