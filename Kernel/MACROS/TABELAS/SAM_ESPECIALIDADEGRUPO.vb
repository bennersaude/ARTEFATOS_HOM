'HASH: 7B019AA59C1F613562C0D4BD2FE8772B
'#Uses "*RegistrarLogAlteracao"

Public Sub TABLE_AfterPost()
  RegistrarLogAlteracao "SAM_ESPECIALIDADEGRUPO", CurrentQuery.FieldByName("HANDLE").AsInteger, "TABLE_AfterPost"
End Sub

Public Sub TABLE_AfterScroll()
  BOTAOGERAREVENTOS.Visible=False
  
  'Luciano T. Alberti - SMS 95290 - 01/04/2008 - Início
  If WebMode Then
    ESPECIALIDADE.ReadOnly = False
    If WebMenuCode = "T1505" Then
      ESPECIALIDADE.ReadOnly = True
    End If
  End If
  'Luciano T. Alberti - SMS 95290 - 01/04/2008 - Fim
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  RegistrarLogAlteracao "SAM_ESPECIALIDADE", CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger, "SAM_ESPECIALIDADEGRUPO.TABLE_BeforeDelete"
End Sub
