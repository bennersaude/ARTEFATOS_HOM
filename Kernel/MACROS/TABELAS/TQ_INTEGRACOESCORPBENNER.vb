'HASH: FA4CB4D60786B5176D36E67D90ADA306

'Macro tabela TQ_INTEGRACOESCORPBENNER

'#Uses "*bsShowMessage"


Public Sub BOTAOINTEGRARDADOS_OnClick()

    Dim vsRetorno As String

    Dim interface As Object
    Set interface =CreateBennerObject("GerenciadorTabelasBasicas.Rotinas")
    vsRetorno = interface.IntegrarDadosGerados(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

    bsShowMessage(vsRetorno,"I")

    If VisibleMode Then
      RefreshNodesWithTable("TQ_INTEGRACOESCORPBENNER")
    End If

End Sub



Public Sub BOTAOPROCESSARDADOS_OnClick()

    Dim vsRetorno As String

    Dim interface As Object
    Set interface =CreateBennerObject("GerenciadorTabelasBasicas.Rotinas")
    vsRetorno = interface.GerarDadosIntegracao(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

    bsShowMessage(vsRetorno,"I")

    If VisibleMode Then
      RefreshNodesWithTable("TQ_INTEGRACOESCORPBENNER")
    End If

End Sub


