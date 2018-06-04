'HASH: 713ADCAD41AB24E66EF3D1C66A13D731
'SAM_PROCREF_PRESTADOR_ITEM_RES
'#Uses "*bsShowMessage"

Option Explicit

Dim vMarca As String
Dim vPontuacao As Double

Public Sub TABLE_AfterScroll()

  If CurrentQuery.FieldByName("DESCRITIVA").AsInteger = 2 Then
    DESCRITIVA.Visible = False
  Else
    DESCRITIVA.Visible = True
  End If

  If CurrentQuery.FieldByName("PERMITEALTERARPONTO").AsString = "N" Then
    PONTOSPADRAO.Visible = False
    PONTOSPADRAOORIGINAL.Visible = False
  Else
    PONTOSPADRAO.Visible = True
    PONTOSPADRAOORIGINAL.Visible = True
  End If


End Sub


Public Sub TABLE_afterEdit()
  vMarca = CurrentQuery.FieldByName("MARCA").AsString
  vPontuacao = CurrentQuery.FieldByName("PONTOSPADRAO").AsFloat
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim q1 As Object
  Dim q2 As Object
  Dim q3 As Object
  Dim qAux As Object

  Dim vPontos As Integer
  Dim qExec As Object

  If vMarca = "N" And CurrentQuery.FieldByName("MARCA").AsString = "N" Then
    If CurrentQuery.FieldByName("PONTOSPADRAO").AsFloat <> vPontuacao Then
      bsShowMessage("Só altere o valor da questão, se mesma for marcada.", "I")
      CanContinue = False
      Exit Sub
    End If
  End If

  If vMarca = "S" And CurrentQuery.FieldByName("MARCA").AsString = "N" Then
    If CurrentQuery.FieldByName("PONTOSPADRAO").AsFloat <> vPontuacao Then
      bsShowMessage("Para desmarcar a questão, não altere o valor da pontuação.", "I")
      CanContinue = False
      Exit Sub
    End If
  End If


  Set q1 = NewQuery
  q1.Add("SELECT HANDLE, TIPO, PONTUACAOMAXIMA, PONTOS FROM SAM_PROCREF_PRESTADOR_AVAL_ITE WHERE HANDLE = :HANDLE")
  q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PROCPRESTAVALITEM").Value
  q1.Active = True

  Set q2 = NewQuery
  q2.Add("SELECT B.TOTAL,B.PONTOS,A.FINALIZACAODATA")
  q2.Add("  FROM SAM_PROCREF_PRESTADOR A")
  q2.Add("  JOIN SAM_PROCREF_PRESTADOR_AVAL B ON (A.HANDLE = B.PROCREFPRESTADOR)")
  q2.Add(" WHERE B.HANDLE = :HANDLE")
  q2.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PROCREF_PRESTADOR_AVAL")
  q2.Active = True

  If Not q2.FieldByName("FINALIZACAODATA").IsNull Then
    bsShowMessage("Este processo de avaliação encontra-se FINALIZADO !", "E")
    CanContinue = False
    Exit Sub
  End If


  'verificar se está reprovado em outro avaliação que não permite continuar processo
  Set q3 = NewQuery
  q3.Active = False
  q3.Clear
  q3.Add("SELECT COUNT(HANDLE) NREC FROM SAM_PROCREF_PRESTADOR_AVAL ")
  q3.Add(" WHERE HANDLE <> :HANDLE AND PROCREFPRESTADOR = :PROCREFPRESTADOR AND SITUACAO = 'R' AND REPROVACAO = 'N'")
  q3.ParamByName("PROCREFPRESTADOR").Value = RecordHandleOfTable("SAM_PROCREF_PRESTADOR")
  q3.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PROCREF_PRESTADOR_AVAL")
  q3.Active = True

  If q3.FieldByName("NREC").AsInteger > 0 Then
    bsShowMessage("Este prestador está reprovado em outra avaliação que não " + Chr(13) + "permite continuar processo quando isto ocorre !", "E")
    CanContinue = False
    Exit Sub
  End If



  If CurrentQuery.FieldByName("PONTOSPADRAO").AsInteger > q1.FieldByName("PONTUACAOMAXIMA").AsInteger Then
    bsShowMessage("Qtde de pontos alterada é maior que a pontuação máxima do item!", "E")
    q1.Active = False
    CanContinue = False
    Exit Sub
  End If

  Set qAux = NewQuery
  If CurrentQuery.FieldByName("MARCA").AsString = "S" Then
    If q1.FieldByName("TIPO").AsInteger = 1 Then
      qAux.Add("SELECT HANDLE FROM SAM_PROCREF_PRESTADOR_ITEM_RES                                   ")
      qAux.Add("WHERE HANDLE <> :HANDLE AND PROCPRESTAVALITEM = :PROCPRESTAVALITEM AND MARCA = 'S' ")
      qAux.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
      qAux.ParamByName("PROCPRESTAVALITEM").Value = CurrentQuery.FieldByName("PROCPRESTAVALITEM").Value
      qAux.Active = True
      If Not qAux.EOF Then
        bsShowMessage("Este item é de resposta única e já existe outra resposta marcada !", "E")
        CanContinue = False
        Exit Sub
      End If
    End If

    'If q1.FieldByName("PONTOS").AsInteger + _
    '   (CurrentQuery.FieldByName("PONTOSPADRAO").AsInteger - CurrentQuery.FieldByName("PONTOS").AsInteger) > _
    '   q1.FieldByName("PONTUACAOMAXIMA").AsInteger Then
    '  If VisibleMode Then
    '    MsgBox "Excedeu o limite da pontuação máxima do item. Operação cancelada !"
    '  End If
    '  CanContinue=False
    '  Exit Sub
    'End If

    If CurrentQuery.FieldByName("DESCRITIVARESPOSTA").IsNull And CurrentQuery.FieldByName("DESCRITIVAOBRIGATORIA").AsString = "S" Then
      bsShowMessage("Necessário informar o campo 'Descritiva resposta' !", "E")
      CanContinue = False
      Exit Sub
    End If


    vPontos = q1.FieldByName("PONTOS").AsInteger + (CurrentQuery.FieldByName("PONTOSPADRAO").AsInteger - CurrentQuery.FieldByName("PONTOS").AsInteger)


    If vPontos > q1.FieldByName("PONTUACAOMAXIMA").AsInteger Then
      vPontos = q1.FieldByName("PONTUACAOMAXIMA").AsInteger
    End If


    Set qExec = NewQuery
    qExec.Add("UPDATE SAM_PROCREF_PRESTADOR_AVAL_ITE                  ")
    qExec.Add("   SET PONTOS = :PONTOS                                ")
    qExec.Add("WHERE HANDLE = :HANDLE                                 ")
    qExec.ParamByName("PONTOS").Value = vPontos
    qExec.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PROCPRESTAVALITEM").Value
    qExec.ExecSQL


    vPontos = (q2.FieldByName("PONTOS").AsInteger - q1.FieldByName("PONTOS").AsInteger) + vPontos

    qExec.Clear
    qExec.Add("UPDATE SAM_PROCREF_PRESTADOR_AVAL         ")
    qExec.Add("   SET PONTOS = :PONTOS,                  ")
    If CurrentQuery.FieldByName("ELIMINATORIA").AsString = "S" Then
      qExec.Add("       SITUACAO = 'R',                    ")
    End If
    qExec.Add("       PERCENTUAL = :PERC                 ")
    qExec.Add("WHERE HANDLE = :HANDLE                    ")
    qExec.ParamByName("PONTOS").Value = vPontos
    qExec.ParamByName("PERC").Value = Round(((vPontos * 100) / q2.FieldByName("TOTAL").AsInteger), 2)
    qExec.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PROCREF_PRESTADOR_AVAL")
    qExec.ExecSQL

    ' If CurrentQuery.FieldByName("ELIMINATORIA").AsString = "S" Then
    '  qExec.Clear
    'qExec.Add("UPDATE SAM_PROCREF_PRESTADOR         ")
    'qExec.Add("   SET SITUACAO = 'R',               ")
    'qExec.Add("       FINALIZACAODATA = :DATA,      ")
    'qExec.Add("       FINALIZACAOUSUARIO = :USUARIO ")
    'qExec.Add("WHERE HANDLE = :HANDLE               ")
    'qExec.ParamByName("DATA").Value = ServerNow
    'qExec.ParamByName("USUARIO").Value = CurrentUser
    'qExec.ParamByName("HANDLE").Value       = RecordHandleOfTable("SAM_PROCREF_PRESTADOR")
    'qExec.ExecSQL
    '; Else
    qAux.Active = False
    qAux.Clear
    qAux.Add("SELECT DISTINCT AV.APROVACAO, A.PERCENTUAL                                  ")
    qAux.Add("  FROM SAM_AVALIACAOREF               AV,                                   ")
    qAux.Add("       SAM_PROCREF_PRESTADOR_AVAL     A                                     ")
    qAux.Add(" WHERE A.AVALIACAOREF = AV.HANDLE                                           ")
    qAux.Add("   AND A.HANDLE = :HANDLERES                                                ")
    qAux.Add("   AND A.PROCREFPRESTADOR = :PROCREFPRESTADOR                               ")
    qAux.Add("   AND 0 = (SELECT COUNT(X.HANDLE)                                          ")
    qAux.Add("               FROM SAM_PROCREF_PRESTADOR_AVAL_ITE I,                       ")
    qAux.Add("                    SAM_PROCREF_PRESTADOR_ITEM_RES X                        ")
    qAux.Add("              WHERE X.PROCPRESTAVALITEM = I.HANDLE                          ")
    qAux.Add("                AND I.PROCESSOREFPRESTAVAL = A.HANDLE                       ")
    qAux.Add("                AND I.HANDLE <> :HANDLEITEM                                 ")
    qAux.Add("                AND 'S' NOT IN (SELECT X.MARCA                              ")
    qAux.Add("                                  FROM SAM_PROCREF_PRESTADOR_ITEM_RES X     ")
    qAux.Add("                                 WHERE X.PROCPRESTAVALITEM = I.HANDLE       ")
    qAux.Add("                               )                                            ")
    qAux.Add("            )                                                               ")
    qAux.ParamByName("HANDLERES").Value = RecordHandleOfTable("SAM_PROCREF_PRESTADOR_AVAL")
    qAux.ParamByName("PROCREFPRESTADOR").Value = RecordHandleOfTable("SAM_PROCREF_PRESTADOR")
    qAux.ParamByName("HANDLEITEM").Value = q1.FieldByName("HANDLE").Value
    qAux.Active = True

    If Not qAux.FieldByName("APROVACAO").IsNull Then
      qExec.Clear
      qExec.Add("UPDATE SAM_PROCREF_PRESTADOR_AVAL      ")
      If qAux.FieldByName("PERCENTUAL").AsCurrency >= qAux.FieldByName("APROVACAO").AsCurrency Then
        qExec.Add("   SET SITUACAO = 'A'                ")
      Else
        qExec.Add("   SET SITUACAO = 'R'                ")
      End If
      qExec.Add("WHERE HANDLE = :HANDLE                 ")
      qExec.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PROCREF_PRESTADOR_AVAL")
      qExec.ExecSQL
    End If
  End If
  CurrentQuery.FieldByName("PONTOS").Value = CurrentQuery.FieldByName("PONTOSPADRAO").Value
  '  End If

  If CurrentQuery.FieldByName("MARCA").AsString = "N" And vMarca = "S" Then
    qAux.Active = False
    qAux.Clear
    qAux.Add("SELECT SUM(PONTOS) TOTAL FROM SAM_PROCREF_PRESTADOR_ITEM_RES WHERE PROCPRESTAVALITEM = :PROCPRESTAVALITEM ")
    qAux.ParamByName("PROCPRESTAVALITEM").Value = CurrentQuery.FieldByName("PROCPRESTAVALITEM").Value
    qAux.Active = True

    If (qAux.FieldByName("TOTAL").AsInteger - CurrentQuery.FieldByName("PONTOS").AsInteger) < q1.FieldByName("PONTUACAOMAXIMA").AsInteger Then

      vPontos = qAux.FieldByName("TOTAL").AsInteger - CurrentQuery.FieldByName("PONTOS").AsInteger

      Set qExec = NewQuery
      qExec.Add("UPDATE SAM_PROCREF_PRESTADOR_AVAL_ITE                  ")
      qExec.Add("   SET PONTOS = :PONTOS                                ")
      qExec.Add("WHERE HANDLE = :HANDLE                                 ")
      qExec.ParamByName("PONTOS").Value = vPontos
      qExec.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PROCPRESTAVALITEM").Value
      qExec.ExecSQL

      If qAux.FieldByName("TOTAL").AsInteger <= q1.FieldByName("PONTUACAOMAXIMA").AsInteger Then
        vPontos = q2.FieldByName("PONTOS").AsInteger - CurrentQuery.FieldByName("PONTOS").Value
      Else
        vPontos = q2.FieldByName("PONTOS").AsInteger - q1.FieldByName("PONTUACAOMAXIMA").AsInteger + vPontos
      End If
      qExec.Clear
      qExec.Add("UPDATE SAM_PROCREF_PRESTADOR_AVAL                              ")
      qExec.Add("   SET PONTOS = :PONTOS,                                       ")
      qExec.Add("       PERCENTUAL = :PERC                                      ")
      qExec.Add("WHERE HANDLE = :HANDLE                                         ")

      qExec.ParamByName("PONTOS").Value = vPontos
      qExec.ParamByName("PERC").Value = Round(((vPontos * 100) / q2.FieldByName("TOTAL").AsInteger), 2)
      qExec.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PROCREF_PRESTADOR_AVAL")
      qExec.ExecSQL
    End If
    CurrentQuery.FieldByName("PONTOS").Value = 0

  End If

End Sub

