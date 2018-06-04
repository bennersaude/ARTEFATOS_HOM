'HASH: 372B65A124DF71845D59BF65A6FBF4C0
'MACRO TABELA: SAM_PRESTADOR_PROC_MEM_ESP_GRP
'#Uses "*bsShowMessage"

Option Explicit

Dim vUsuarioResponsavel As Integer

Public Sub ESPECIALIDADEGRUPO_OnPopup(ShowPopup As Boolean)
 UpdateLastUpdate("SAM_ESPECIALIDADEGRUPO")
End Sub

Public Sub Condicao()
 Dim qMEM As Object
 Set qMEM = NewQuery
 Dim vCondicao As String
 Dim vEspecialidade As String
 Dim vPrestador As String

 qMEM.Add("SELECT A.MEMBRO, A.PRESTADOR, B.TABTIPOMOVIMENTACAO, B.ESPECIALIDADE, A.RESPONSAVEL")
 qMEM.Add("  FROM SAM_PRESTADOR_PROC_MEMBROS A     ")
 qMEM.Add("  JOIN SAM_PRESTADOR_PROC_MEM_ESP B ON (A.HANDLE = B.PRESTADORPROCMEMBRO)")
 qMEM.Add(" WHERE B.HANDLE = :HANDLE")
 qMEM.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC_MEM_ESP")
 qMEM.Active = True

 vEspecialidade = qMEM.FieldByName("ESPECIALIDADE").AsString
 vUsuarioResponsavel = qMEM.FieldByName("RESPONSAVEL").AsInteger

 If qMEM.FieldByName("TABTIPOMOVIMENTACAO").AsInteger = 1 Then
         vPrestador = qMEM.FieldByName("PRESTADOR").AsString
 Else
         vPrestador = qMEM.FieldByName("MEMBRO").AsString
 End If

    vCondicao = "A.ESPECIALIDADE = " &  vEspecialidade
    vCondicao = vCondicao + "   AND (EXISTS (SELECT 1                                        "
    vCondicao = vCondicao + "                 FROM SAM_PRESTADOR_ESPECIALIDADEGRP G          "
    vCondicao = vCondicao + "                WHERE A.HANDLE = G.ESPECIALIDADEGRUPO           "
    vCondicao = vCondicao + "                  AND G.ESPECIALIDADE = A.ESPECIALIDADE         "
    vCondicao = vCondicao + "                  AND G.PRESTADOR = " & vPrestador & ")"
    vCondicao = vCondicao + "        OR NOT EXISTS (SELECT 1                                 "
    vCondicao = vCondicao + "                         FROM SAM_PRESTADOR_ESPECIALIDADEGRP B  "
    vCondicao = vCondicao + "                        WHERE B.ESPECIALIDADE = A.ESPECIALIDADE "
    vCondicao = vCondicao + "                          AND B.PRESTADOR = " & vPrestador & "))"


     ESPECIALIDADEGRUPO.LocalWhere = vCondicao

 Set qMEM = Nothing

End Sub

Public Sub TABLE_AfterScroll()
 Condicao
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

    Condicao

    If (EhProcessoFinalizado()) Then
        bsShowMessage("Processo finalizado! Não é possível Inserir!", "E")
        CanContinue = False
        Exit Sub
    End If

    If vUsuarioResponsavel <> CurrentUser Then
          bsShowMessage("Usuário não é o responsável! Inserção não permitida.", "E")
          CanContinue = False
    End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

    If (EhProcessoFinalizado()) Then
        bsShowMessage("Processo finalizado! Não é possível Alterar!", "E")
        CanContinue = False
        Exit Sub
    End If

    If vUsuarioResponsavel <> CurrentUser Then
          bsShowMessage("Usuário não é o responsável! Alteração não permitida.", "E")
          CanContinue = False
    End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

    If EhProcessoFinalizado() Then
        bsShowMessage("Processo finalizado! Não é possível excluir!", "E")
        CanContinue = False
        Exit Sub
    End If

    If vUsuarioResponsavel <> CurrentUser Then
        bsShowMessage("Usuário não é o responsável! Exclusão não permitida.", "E")
        CanContinue = False
    End If
End Sub

Public Function EhProcessoFinalizado As Boolean
 Dim componente As CSBusinessComponent

 Set componente = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcMemEspGrpBLL, Benner.Saude.Prestadores.Business")
 componente.AddParameter(pdtInteger, RecordHandleOfTable("SAM_PRESTADOR_PROC"))
 EhProcessoFinalizado = CBool(componente.Execute("VerificarProcessoFinalizado"))
 Set componente = Nothing
End Function
