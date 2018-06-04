'HASH: F78DEEC95C2DDEBBBD5D0626F775EF18
 'MACRO DA TABELA GRUPOS_ORGANOGRAMA

 '#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  'Balani SMS 71092 06/12/2006
  Dim q As Object
  Set q = NewQuery

  'Verificando se existe outro organograma neste mesmo grupo de seguranca
  q.Add("SELECT HANDLE FROM GRUPOS_ORGANOGRAMA WHERE GRUPOUSUARIOS = :GRUPOUSUARIOS AND ORGANOGRAMA = :ORGANOGRAMA AND HANDLE <> :HANDLE")
  q.ParamByName("GRUPOUSUARIOS").AsInteger = CurrentQuery.FieldByName("GRUPOUSUARIOS").AsInteger
  q.ParamByName("ORGANOGRAMA").AsInteger = CurrentQuery.FieldByName("ORGANOGRAMA").AsInteger
  q.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  q.Active = True

  If Not q.FieldByName("HANDLE").IsNull Then
    bsShowMessage("Organograma já cadastrado para este grupo de segurança.", "E")
    CanContinue = False
    Set q = Nothing
    Exit Sub
  End If


  If CurrentQuery.State = 2 Then 'EDICAO

    'verificar se a alteracao e para remover o flag padrao
    If CurrentQuery.FieldByName("PADRAO").AsString = "N" Then
      q.Active = False
      q.Clear
      q.Add("SELECT PADRAO FROM GRUPOS_ORGANOGRAMA WHERE HANDLE = :HANDLE")
      q.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      q.Active = True

      If q.FieldByName("PADRAO").AsString = "S" Then

        bsShowMessage("Não é permitido alterar o flag padrão, antes é necessário marcar outro organograma deste grupo de segurança como padrão.", "E")
        CanContinue = False
        Set q = Nothing
        Exit Sub

      End If

    End If


    If CurrentQuery.FieldByName("PADRAO").AsString = "S" Then

      'Alterando o antigo padrao para 'N'
      If Not InTransaction Then
        StartTransaction
      End If

      q.Active = False
      q.Clear
      q.Add("UPDATE GRUPOS_ORGANOGRAMA SET PADRAO = 'N' WHERE GRUPOUSUARIOS = :GRUPOUSUARIOS AND HANDLE <> :HANDLE")
      q.ParamByName("GRUPOUSUARIOS").AsInteger = CurrentQuery.FieldByName("GRUPOUSUARIOS").AsInteger
      q.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      q.ExecSQL

      If InTransaction Then
        Commit
      End If

    End If

  Else 'INSERCAO

    If CurrentQuery.FieldByName("PADRAO").AsString = "S" Then

      'Alterando o antigo padrao para 'N'
      If Not InTransaction Then
        StartTransaction
      End If

      q.Active = False
      q.Clear
      q.Add("UPDATE GRUPOS_ORGANOGRAMA SET PADRAO = 'N' WHERE GRUPOUSUARIOS = :GRUPOUSUARIOS")
      q.ParamByName("GRUPOUSUARIOS").AsInteger = CurrentQuery.FieldByName("GRUPOUSUARIOS").AsInteger
      q.ExecSQL

      If InTransaction Then
        Commit
      End If

    Else

      'Verificar se existe algum registro ja gravado como padrao para este grupo
      q.Active = False
      q.Clear
      q.Add("SELECT HANDLE FROM GRUPOS_ORGANOGRAMA WHERE GRUPOUSUARIOS = :GRUPOUSUARIOS AND PADRAO = 'S'")
      q.ParamByName("GRUPOUSUARIOS").AsInteger = CurrentQuery.FieldByName("GRUPOUSUARIOS").AsInteger
      q.Active = True


      If q.FieldByName("HANDLE").IsNull Then

        'Marcar como padrao
        CurrentQuery.FieldByName("PADRAO").AsString = "S"

      End If

    End If

  End If

  Set q = Nothing
  RefreshNodesWithTable("GRUPOS_ORGANOGRAMA")
  'final Balani SMS 71092 06/12/2006
End Sub
