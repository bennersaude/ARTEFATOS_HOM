'HASH: 3BB9493677F0B3AB7FB20FFC21507F76

Public Sub BOTAOGERAR_OnClick()
  Dim Obj As Object


  Set Obj = CreateBennerObject("SamGerarEventos.Rotinas")
  Obj.Gerar(CurrentSystem, CurrentQuery.FieldByName("ITEM").AsInteger)
  Set Obj = Nothing

End Sub


Public Sub BOTAOEXCLUIR_OnClick()
  Dim Obj As Object


  Set Obj = CreateBennerObject("SamGerarEventos.Rotinas")
  Obj.Excluir(CurrentSystem, CurrentQuery.FieldByName("ITEM").AsInteger)
  Set Obj = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim vEventoCorrente As Long
Dim vGrauCorrente As Long
Dim vGrupoBenefCorrente As String

vEventoCorrente = CurrentQuery.FieldByName("EVENTO").AsInteger
vGrauCorrente = CurrentQuery.FieldByName("GRAU").AsInteger
vGrupoBenefCorrente = CurrentQuery.FieldByName("GRUPOBENEF").AsString

  Dim Q As Object
  Set Q = NewQuery
  Q.Clear
  Q.Add("SELECT A.HANDLE ,                                           ")
  Q.Add("       A.EVENTO         EVENTOHANDLE ,                      ")
  Q.Add("       A.GRAU           GRAUHANDLE,                         ")
  Q.Add("       A.GRUPOBENEF     GRUPO                               ")
  Q.Add("  FROM GER_ITEM_EVENTO A                                    ")
  Q.Add("  LEFT JOIN SAM_TGE SAM_TGE ON ( SAM_TGE.HANDLE = A.EVENTO )")
  Q.Add(" WHERE ( A.ITEM = :ITEM )                                   ")
  Q.Add(" ORDER BY SAM_TGE.Z_DESCRICAO ,                             ")
  Q.Add("          A.GRUPOBENEF                                      ")

  Q.ParamByName("ITEM").Value = CurrentQuery.FieldByName("ITEM").AsInteger
  Q.Active = True



  While Not Q.EOF
    If(vEventoCorrente = Q.FieldByName("EVENTOHANDLE").AsInteger And _
       vGrauCorrente = Q.FieldByName("GRAUHANDLE").AsInteger And _
       vGrupoBenefCorrente = Q.FieldByName("GRUPO").AsString) Then

       MsgBox("Este registro não pode ser inserido, já existe um identico.")
       CanContinue = False
       Exit Sub

    End If
    Q.Next
  Wend

  Set Q = Nothing
End Sub
