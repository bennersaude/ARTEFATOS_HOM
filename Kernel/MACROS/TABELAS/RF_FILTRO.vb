'HASH: 8EBB6E50B026B3A513040BF428910F03
'Macro: RF_FILTRO


Public Sub BOTAOIMPRIMIR_OnClick()

  Dim Filtro As String

  If CurrentQuery.FieldByName("RELATORIO").AsString = "1" Then 'pRESTADOR

    Filtro = "xxxxxxxxxxxxxx"




    ReportPrint 128, Filtro, False, False
  End If

  If CurrentQuery.FieldByName("RELATORIO").AsString = "1" Then 'pRESTADOR
  End If




End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)

  Dim interface As Object
  Dim vHandle As Long


  ShowPopup = False

  Set interface = CreateBennerObject("Procura.Procurar")


  vHandle = interface.Exec(CurrentSystem, "SAM_PRESTADOR", "NOME|PRESTADOR", 1, "NOME|CPNF/CNPJ", "", "PRESTADOR", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If
  Set INTERFACE = Nothing
End Sub

