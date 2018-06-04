'HASH: 7E0557D24E3974CE5187A744AF1CC255
'#Uses "*bsShowMessage"
'Macro: SAM_ROTINAIMPCONVRECIP_DADOS

Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)
  ShowPopup = False
  Dim vHandle As Long
  Dim Interface As Object
  Dim vColunas, vCriterio, vCampos, vTabela, vData As String
  Dim qMatricula As Object


  ShowPopup = False

  vData = SQLDate(ServerDate)

  Set Interface = CreateBennerObject("Procura.Procurar")

  vTabela = "SAM_BENEFICIARIO"
  vColunas = "BENEFICIARIO|NOME|DATAADESAO|DATACANCELAMENTO"
  vCriterio = "(SAM_BENEFICIARIO.DATACANCELAMENTO IS NULL OR SAM_BENEFICIARIO.DATACANCELAMENTO > "+vData+")"
  vCriterio = vCriterio + _
              " AND NOT EXISTS(SELECT 1 "+ _
              "                  FROM SAM_ROTINAIMPCONVRECIP_DADOS X"+ _
              "                 WHERE X.ROTINAIMP = "+ CStr(RecordHandleOfTable("SAM_ROTINAIMPCONVRECIP")) + _
              "                   AND X.BENEFICIARIO = SAM_BENEFICIARIO.HANDLE)

  vCampos = "Beneficiário|Nome|Data adesão|Data cancelamento"

  vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 2, vCampos, vCriterio, "Tabela de beneficiários", False, "")

  If vHandle <>0 Then
    Set qMatricula = NewQuery
    qMatricula.Add("SELECT B.NOME, M.NOMEMAE, M.NOMEPAI, M.SEXO, M.DATANASCIMENTO, M.CARTAONACIONALSAUDE, M.CPF ")
    qMatricula.Add("  FROM SAM_BENEFICIARIO B ")
    qMatricula.Add("  JOIN SAM_MATRICULA    M ON (M.HANDLE = B.MATRICULA) ")
    qMatricula.Add(" WHERE B.HANDLE = :HANDLE ")
    qMatricula.ParamByName("HANDLE").AsInteger = vHandle
    qMatricula.Active=True
    CurrentQuery.Edit
    CurrentQuery.FieldByName("BENEFICIARIO").Value = vHandle
    CurrentQuery.FieldByName("NOME").Value = qMatricula.FieldByName("NOME").Value
    CurrentQuery.FieldByName("NOMEMAE").Value = qMatricula.FieldByName("NOMEMAE").Value
    CurrentQuery.FieldByName("NOMEPAI").Value = qMatricula.FieldByName("NOMEPAI").Value
    CurrentQuery.FieldByName("SEXO").Value = qMatricula.FieldByName("SEXO").Value
    CurrentQuery.FieldByName("DATANASCIMENTO").Value = qMatricula.FieldByName("DATANASCIMENTO").Value
    CurrentQuery.FieldByName("CARTAONACIONALSAUDE").Value = qMatricula.FieldByName("CARTAONACIONALSAUDE").Value
    CurrentQuery.FieldByName("CPF").Value = qMatricula.FieldByName("CPF").Value
    Set qMatricula = Nothing
  End If

  Set Interface = Nothing

End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  ShowPopup = False
  Dim vHandle As Long
  Dim Interface As Object
  Dim vColunas, vCriterio, vCampos, vTabela, vData As String


  ShowPopup = False

  Set Interface = CreateBennerObject("Procura.Procurar")

  vTabela = "SAM_PRESTADOR"
  vColunas = "PRESTADOR|NOME"
  vCriterio = "CONVENIORECIPROCIDADE = 'S'"

  vCampos = "Prestador|Nome"

  vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 2, vCampos, vCriterio, "Tabela de prestadores", False, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If

  Set Interface = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  Dim Q As Object
  Set Q = NewQuery
  Q.Add("SELECT USUARIOPROCESSO FROM SAM_ROTINAIMPCONVRECIP WHERE HANDLE = :HANDLE")
  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("ROTINAIMP").AsInteger
  Q.Active=True
  If Not Q.FieldByName("USUARIOPROCESSO").IsNull Then
    RecordReadOnly = True
  Else
    RecordReadOnly = False
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("PRESTADOR").IsNull Then
    bsShowMessage("Campo Prestador é obrigatório.","E")
    CanContinue=False
    Exit Sub
  End If
End Sub
