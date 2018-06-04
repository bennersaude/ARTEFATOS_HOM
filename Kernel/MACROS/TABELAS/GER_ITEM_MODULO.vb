'HASH: 8FA7BDE1A8346E5A6C54D0C3EF1AE590

Public Sub BOTAOGERAR_OnClick()
  Dim Obj As Object


  Set Obj = CreateBennerObject("SamGerarExcluirDados.Rotinas")
  Obj.Gerar(CurrentSystem, "GER_ITEM_MODULO", "Duplicando Modulos para o Item", "SAM_MODULO", "MODULO", "ITEM", CurrentQuery.FieldByName("ITEM").AsInteger, "N", "TipoProduto")
  Set Obj = Nothing

End Sub


Public Sub BOTAOEXCLUIR_OnClick()
  Dim Obj As Object


  Set Obj = CreateBennerObject("SamGerarExcluirDados.Rotinas")
  Obj.Excluir(CurrentSystem, "GER_ITEM_MODULO", "Excluindo Modulo do Item", "SAM_MODULO", "MODULO", "ITEM", CurrentQuery.FieldByName("ITEM").AsInteger, "N", "TipoProduto")
  Set Obj = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim q As Object
  Set q = NewQuery
  q.Clear

  q.Add("SELECT 1 FROM ger_item_modulo WHERE modulo=:modulo AND item=:item ")

  q.ParamByName("modulo").Value = CurrentQuery.FieldByName("modulo").Value
  q.ParamByName("item"  ).Value = CurrentQuery.FieldByName("item"  ).Value

  q.Active = True

  If Not q.EOF Then
    CanContinue = False
   	MsgBox("Este módulo já está cadastrado para este item ! ")

  End If
End Sub
