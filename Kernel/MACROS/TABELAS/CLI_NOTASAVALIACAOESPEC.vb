'HASH: 815FEACBC96A19D111DF283814C06818
Option Explicit
'#USES "*PrimeiroDiaCompetencia"
'#USES "*UltimoDiaCompetencia"

Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vWhere As String
  Dim vColunas As String
  Dim vCabecalho As String
  Dim vTabela As String
  Dim vPrimeiroDia, vUltimoDia As Date
  Dim vRecurso As String
  Dim vHandle As Long
  Dim sql As Object
  Set sql = NewQuery

  ShowPopup = False

  sql.Clear
  sql.Add("SELECT COMPETENCIA FROM CLI_NOTASAVALIACAO")
  sql.Add(" WHERE HANDLE = :ROTINA")
  sql.ParamByName("ROTINA").AsInteger = RecordHandleOfTable("CLI_NOTASAVALIACAO")
  sql.Active = True

  vPrimeiroDia = PRIMEIRODIACOMPETENCIA(sql.FieldByName("COMPETENCIA").AsDateTime)
  vUltimoDia = ULTIMODIACOMPETENCIA(sql.FieldByName("COMPETENCIA").AsDateTime)

  sql.Clear
  sql.Add("SELECT RECURSO FROM CLI_NOTASAVALIACAORECURSO")
  sql.Add(" WHERE HANDLE = :ROTINA")
  sql.ParamByName("ROTINA").AsInteger = RecordHandleOfTable("CLI_NOTASAVALIACAORECURSO")
  sql.Active = True

  vRecurso = sql.FieldByName("RECURSO").AsString

  vCabecalho = "Nome|Beneficiario"
  vColunas = "SAM_MATRICULA.Z_NOME|SAM_BENEFICIARIO.BENEFICIARIO"
  vTabela = "SAM_BENEFICIARIO|SAM_MATRICULA[SAM_BENEFICIARIO.MATRICULA = SAM_MATRICULA.HANDLE]"
  vWhere = "EXISTS (SELECT 1 FROM CLI_AGENDA A, CLI_ATENDIMENTO T "
  vWhere = vWhere + "WHERE T.AGENDA = A.HANDLE"
  vWhere = vWhere + "  AND A.DATAMARCADA BETWEEN " + SQLDate(vPrimeiroDia) + " AND " + SQLDate(vUltimoDia)
  vWhere = vWhere + "  AND T.HORAFINAL IS NOT  NULL"
  vWhere = vWhere + "  AND A.RECURSO = " + vRecurso
  vWhere = vWhere + "  AND A.BENEFICIARIO = SAM_BENEFICIARIO.HANDLE)"

  Set interface = CreateBennerObject("Procura.Procurar")
  vHandle = interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCabecalho, vWhere, "Procura por beneficiário", True, "")

  If vHandle >0 Then
    CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = vHandle
  End If
  Set interface = Nothing
  Set sql = Nothing
End Sub

