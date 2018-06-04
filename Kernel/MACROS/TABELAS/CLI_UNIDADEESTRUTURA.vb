'HASH: 6574FAE9FA6D767F0019CE9CD94D5894

Option Explicit

'#Uses "*bsShowMessage"

Dim DataAteNula         As Boolean
Dim HierarquiaAnterior  As Long

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
    If Not PermitirServicoProprio Then
      CanContinue = False
      bsshowmessage("Não é possível excluir essa unidade. Favor verificar se ela possui alguma unidade dependente ou se possui profissional de saúde vinculado.","E")
    End If
End Sub

Public Function PermitirExclusao As Boolean

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT 1                    ")
  SQL.Add("  FROM CLI_UNIDADEESTRUTURA ")
  SQL.Add(" WHERE HIERARQUIA = :HANDLE ")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  If Not SQL.EOF Then
    PermitirExclusao = False
  Else
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT 1                            ")
    SQL.Add("  FROM CLI_RECURSO_UNIDADEESTRUTURA ")
    SQL.Add(" WHERE UNIDADE = :HANDLE            ")
    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.Active = True

    If Not SQL.EOF Then
      PermitirExclusao = False
    Else
      PermitirExclusao = True
    End If

  End If

End Function


Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DataAteNula = True
  Else
    DataAteNula = False
  End If

  HierarquiaAnterior = CurrentQuery.FieldByName("HIERARQUIA").AsInteger

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  'A verificação não poderá ser feita na inclusão
  If CurrentQuery.State <> 3 Then
    If Not CurrentQuery.FieldByName("DATAFINAL").IsNull And DataAteNula Then 'A data final era nula e agora possui valor
      If Not PermitirDataAte Then
        CanContinue = False
        bsshowmessage("Não é possível incluir uma data até se a unidade da estrutura possui unidades dependentes.","E")
      End If
    End If
  End If

  If Not CurrentQuery.FieldByName("SERVICOPROPRIO").IsNull Then 'Verificar se o serviço próprio informado já não está vinculado a outra unidade
    If Not PermitirServicoProprio Then
      CanContinue = False
      bsshowmessage("O Serviço Próprio que está tentando selecionar já está vinculado a outra Unidade da Estrutura.","E")
    End If
  End If

  If Not CurrentQuery.FieldByName("UNIDADE").IsNull Then 'Verificar se a filial padrão informada já não está vinculada a outra unidade
    If Not PermitirFilialPadrao Then
      CanContinue = False
      bsshowmessage("A filial padrão selecionada já está vinculada a outra unidade da estrutura! Favor verificar.","E")
    End If
  End If

  'A verificação não poderá ser feita na inclusão, pois ainda não tem nenhum filho
  If CurrentQuery.State <> 3 Then
    If (HierarquiaAnterior <> CurrentQuery.FieldByName("HIERARQUIA").AsInteger) And (Not CurrentQuery.FieldByName("HIERARQUIA").IsNull) Then 'A hierarquia atual é diferente da anterior
      If Not PermitirHierarquia Then
        CanContinue = False
        bsshowmessage("Não é possível atribuir a hierarquia desejada, pois a mesma é uma unidade dependente.","E")
      End If
    End If
  End If

End Sub

Public Function PermitirHierarquia As Boolean

  Dim vsUnidadesDependentes As String
  vsUnidadesDependentes = ListarUnidadesDependentes(CurrentQuery.FieldByName("HANDLE").AsInteger)

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT 1                                             ")
  SQL.Add("  FROM DUAL                                          ")
  SQL.Add(" WHERE :HIERARQUIA IN (" + vsUnidadesDependentes + ")")
  SQL.ParamByName("HIERARQUIA").AsInteger = CurrentQuery.FieldByName("HIERARQUIA").AsInteger
  SQL.Active = True

  If SQL.EOF Then
    PermitirHierarquia = True
  Else
    PermitirHierarquia = False
  End If

  SQL.Active = False
  Set SQL = Nothing

End Function

Public Function ListarUnidadesDependentes(piHandle As Long) As String

  ListarUnidadesDependentes = "0"

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT HANDLE               ")
  SQL.Add("  FROM CLI_UNIDADEESTRUTURA ")
  SQL.Add(" WHERE HIERARQUIA = :HANDLE ")
  SQL.ParamByName("HANDLE").AsInteger = piHandle
  SQL.Active = True

  While Not SQL.EOF
    If ListarUnidadesDependentes = "0" Then
      ListarUnidadesDependentes = SQL.FieldByName("HANDLE").AsString + "," + ListarUnidadesDependentes(SQL.FieldByName("HANDLE").AsInteger)
    Else
      ListarUnidadesDependentes = ListarUnidadesDependentes + "," + SQL.FieldByName("HANDLE").AsString + "," + ListarUnidadesDependentes(SQL.FieldByName("HANDLE").AsInteger)
    End If
    SQL.Next
  Wend

  SQL.Active = False
  Set SQL = Nothing

End Function

Public Function PermitirDataAte As Boolean

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT 1                                                      ")
  SQL.Add("  FROM CLI_UNIDADEESTRUTURA                                   ")
  SQL.Add(" WHERE HIERARQUIA = :HANDLE                                   ")
  SQL.Add("   AND (DATAFINAL IS NULL OR DATAFINAL >= :DATAFINALCORRENTE) ")
  SQL.ParamByName("HANDLE").AsInteger             = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("DATAFINALCORRENTE").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
  SQL.Active = True

  If SQL.EOF Then
    PermitirDataAte = True
  Else
    PermitirDataAte = False
  End If

  SQL.Active = False
  Set SQL = Nothing

End Function

Public Function PermitirServicoProprio As Boolean

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT 1                                ")
  SQL.Add("  FROM CLI_UNIDADEESTRUTURA             ")
  SQL.Add(" WHERE SERVICOPROPRIO = :SERVICOPROPRIO ")
  SQL.Add("   AND HANDLE        <> :HANDLE         ")
  SQL.ParamByName("HANDLE").AsInteger         = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("SERVICOPROPRIO").AsInteger = CurrentQuery.FieldByName("SERVICOPROPRIO").AsInteger
  SQL.Active = True

  If SQL.EOF Then
    PermitirServicoProprio = True
  Else
    PermitirServicoProprio = False
  End If

  SQL.Active = False
  Set SQL = Nothing

End Function

Public Function PermitirFilialPadrao As Boolean

  Dim viHandlePai As Long
  viHandlePai = ObterNoPai

  Dim vsUnidadesPai As String
  vsUnidadesPai = ListarUnidadesDependentes(viHandlePai) + "," + Str(viHandlePai)

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT 1                                    ")
  SQL.Add("  FROM CLI_UNIDADEESTRUTURA                 ")
  SQL.Add(" WHERE UNIDADE = :UNIDADE                   ")
  SQL.Add("   AND HANDLE <> :HANDLE                    ")
  SQL.Add("   AND HANDLE NOT IN (" + vsUnidadesPai + ")")
  SQL.ParamByName("HANDLE").AsInteger  = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("UNIDADE").AsInteger = CurrentQuery.FieldByName("UNIDADE").AsInteger
  SQL.Active = True

  If SQL.EOF Then
    PermitirFilialPadrao = True
  Else
    PermitirFilialPadrao = False
  End If

  SQL.Active = False
  Set SQL = Nothing

End Function

Public Function ObterNoPai As Long

  Dim viHAux As Long
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT HANDLE,              ")
  SQL.Add("       HIERARQUIA           ")
  SQL.Add("  FROM CLI_UNIDADEESTRUTURA ")
  SQL.Add(" WHERE HANDLE = :HANDLE     ")

  If Not CurrentQuery.FieldByName("HIERARQUIA").IsNull Then

    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.Active = True

    While Not SQL.FieldByName("HIERARQUIA").IsNull

      viHAux = SQL.FieldByName("HIERARQUIA").AsInteger
      SQL.Active = False
      SQL.ParamByName("HANDLE").AsInteger  = viHAux
      SQL.Active = True

    Wend

    ObterNoPai =  SQL.FieldByName("HANDLE").AsInteger

  Else

    ObterNoPai = CurrentQuery.FieldByName("HANDLE").AsInteger

  End If

  SQL.Active = False
  Set SQL = Nothing

End Function

