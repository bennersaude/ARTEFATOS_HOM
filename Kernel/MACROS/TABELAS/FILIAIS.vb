'HASH: E59CB52800BE2291B9A778D4570B9F2B
'#Uses "*bsShowMessage"

Public Sub FILIALPROCESSAMENTO_OnPopup(ShowPopup As Boolean)
  UpdateLastUpdate("FILIAIS")
End Sub

Public Sub TABLE_AfterPost()
  If Not VisibleMode Then
    Exit Sub
  End If


  If CurrentQuery.FieldByName("FILIALPROCESSAMENTO").IsNull Then

    Dim SQL As Object

    Set SQL = NewQuery

    SQL.Clear
    SQL.Add("UPDATE FILIAIS SET FILIALPROCESSAMENTO = :FILIALPROCESSAMENTO")
    SQL.Add(" WHERE HANDLE = :HANDLE")
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ParamByName("FILIALPROCESSAMENTO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ExecSQL

    Set SQL = Nothing

    CurrentQuery.Active = False
    CurrentQuery.Active = True

  End If

  Dim TQIntegracoesCorpBennerBLL As CSBusinessComponent
  Set TQIntegracoesCorpBennerBLL = BusinessComponent.CreateInstance("Benner.Saude.IntegracaoFinanceira.Business.TabelasBasicas.TQIntegracoesCorpBennerBLL, Benner.Saude.IntegracaoFinanceira.Business")

  TQIntegracoesCorpBennerBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "FILIAIS")
  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "X")

  TQIntegracoesCorpBennerBLL.Execute("InserirDadosIntegracao")
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If Not VisibleMode Then
    Exit Sub
  End If

  If CurrentQuery.State = 2 Then
    If CurrentQuery.FieldByName("FILIALPROCESSAMENTO").IsNull Then
      bsShowMessage("Preencha Filial de Processamento.", "E")
      CanContinue = False
    End If
  End If

End Sub
