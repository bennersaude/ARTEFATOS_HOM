'HASH: 93BD6A5B3675C34C1D2E49BA6F9C5C67
'#Uses "*bsShowMessage"
' Tabela : SAM_MATMED

Dim vCompra As Boolean

Option Explicit

Public Sub CADASTROCOMPRA_OnChange()
  If ((CurrentQuery.State = 3) Or (CurrentQuery.State = 2)) Then
    vCompra = Not vCompra
    If (vCompra) Then
      BRFORNECEDOR.ReadOnly = True
      BRFORNECEDOR.Required = False
      FORNECEDORPESSOA.ReadOnly = False
      FORNECEDORPESSOA.Required = True
      CurrentQuery.FieldByName("TABORIGEM").AsInteger = 1
      CurrentQuery.FieldByName("TABTIPO").AsInteger = 1
      CurrentQuery.FieldByName("BRFORNECEDOR").Value = Null
      TABORIGEM.ReadOnly = True
      TABTIPO.ReadOnly = True
    Else
      BRFORNECEDOR.ReadOnly = False
      BRFORNECEDOR.Required = True
      FORNECEDORPESSOA.ReadOnly = True
      FORNECEDORPESSOA.Required = False
      TABORIGEM.ReadOnly = False
      TABTIPO.ReadOnly = False
      CurrentQuery.FieldByName("FORNECEDORPESSOA").Value = Null
    End If
  End If
End Sub

Public Sub GENERICO_OnChange()
  Dim SQL As Object
  Dim matmed As String


  If CurrentQuery.FieldByName("GENERICO").AsString = "S" Then
    Set SQL = NewQuery

    SQL.Add("SELECT DESCRICAO FROM SAM_MATMED WHERE CODIGOGENERICO = " + CurrentQuery.FieldByName("HANDLE").AsString)
    SQL.Active = True

    If Not SQL.EOF Then
      While Not SQL.EOF
        matmed = matmed + "- " + SQL.FieldByName("DESCRICAO").AsString
        SQL.Next
      Wend
      bsShowMessage("Existe um medicamento apontando para este genérico", "I")
      CurrentQuery.FieldByName("generico").AsString = "N"
    End If

    Set SQL = Nothing

  End If
End Sub

Public Sub TABLE_AfterInsert()
  vCompra = False
End Sub

Public Sub TABLE_AfterScroll()
  If (CurrentQuery.FieldByName("CADASTROCOMPRA").AsString = "S") Then
    BRFORNECEDOR.ReadOnly = True
    FORNECEDORPESSOA.ReadOnly = False
    TABORIGEM.ReadOnly = True
    TABTIPO.ReadOnly = True
    vCompra = True
  Else
    BRFORNECEDOR.ReadOnly = False
    FORNECEDORPESSOA.ReadOnly = True
    TABORIGEM.ReadOnly = False
    TABTIPO.ReadOnly = False
    vCompra = False
  End If

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If (CurrentQuery.FieldByName("CADASTROCOMPRA").AsString = "S") Then
    vCompra = True
  Else
    vCompra = False
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If (CurrentQuery.FieldByName("GENERICO").AsString = "S") And Not(CurrentQuery.FieldByName("CODIGOGENERICO").IsNull) Then
    bsShowMessage("O campo 'Genérico deste medicamente' será desmarcado, pois este está marcado como genérico", "I")
    CurrentQuery.FieldByName("CODIGOGENERICO").Clear
  End If

  If ((FORNECEDORPESSOA.ReadOnly = False) And (CurrentQuery.FieldByName("FORNECEDORPESSOA").IsNull)) Then
    bsShowMessage("Necessário informar 'Fornecedor para compra'", "E")
    FORNECEDORPESSOA.SetFocus
    CanContinue = False
  End If

End Sub

