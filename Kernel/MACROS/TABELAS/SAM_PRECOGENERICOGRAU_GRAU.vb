'HASH: D25DA3DA17FEA389E1E99FB4366F8C42

'Macro: SAM_PRECOGENERICO_GRAU

Option Explicit

Dim gTipoGrau As String


'#Uses "*ProcuraEvento"

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)

  '  If Len(EVENTO.Text)=0 Then
  Dim vHandle As Long
  Dim INTERFACE As Object
  Dim vColunas, vCriterio, vCampos, vTabela As String

  Set INTERFACE = CreateBennerObject("Procura.Procurar")

  ShowPopup = False
  If CurrentQuery.FieldByName("GRAU").IsNull Then
    MsgBox("Escolha o grau primeiro")
    'CanContinue=False
    Exit Sub
  End If
  'vHandle =ProcuraEvento(True,EVENTO.Text)
  vColunas = "SAM_TGE.ESTRUTURA|SAM_TGE.DESCRICAO"
  vCriterio = "(SAM_TGE_GRAU.GRAU =  " + CurrentQuery.FieldByName("GRAU").AsString + "  )"
  vCampos = "Código do evento|Descrição"
  vTabela = "SAM_TGE|SAM_TGE_GRAU[SAM_TGE_GRAU.EVENTO = SAM_TGE.HANDLE]"

  vHandle = INTERFACE.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, vCriterio, "Eventos que tem como grau válido o grau acima", True, EVENTO.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If
  '  End If
End Sub
Public Sub GRAU_OnChange()
  CurrentQuery.FieldByName("EVENTO").Clear
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DATAFINAL.ReadOnly = False
  Else
    DATAFINAL.ReadOnly = True
  End If

  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT TIPOGRAU FROM SAM_PRECOGENERICOGRAU WHERE HANDLE = " + Str(RecordHandleOfTable("SAM_PRECOGENERICOGRAU")))
  SQL.Active = True
  gTipoGrau = SQL.FieldByName("TIPOGRAU").AsString
  If gTipoGrau = "P" Then 'pacote
    GRAU.LocalWhere = "ORIGEMVALOR='7'"
  Else
    GRAU.LocalWhere = "ORIGEMVALOR<>'7' AND (PRECOPORGRAU='S' OR PRECOPORGRAUDOTACAO='S')"
  End If
  Set SQL = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim INTERFACE As Object
  Dim QueryGrau As Object
  Dim Linha As String
  Dim Condicao As String
  Dim HGrau As Long
  Dim NGrau As Long

  If(gTipoGrau = "P")And(CurrentQuery.FieldByName("EVENTO").IsNull)Then
  MsgBox "É necessário informar o evento"
  CanContinue = False
  Exit Sub
End If


Set QueryGrau = NewQuery
Set INTERFACE =CreateBennerObject("SAMGERAL.Vigencia")
HGrau = CurrentQuery.FieldByName("HANDLE").AsInteger
NGrau = CurrentQuery.FieldByName("GRAU").AsInteger

If HGrau <> -1 Then
  QueryGrau.Add("SELECT HANDLE,")
  QueryGrau.Add("       DATAFINAL,")
  QueryGrau.Add("       GRAU,")
  QueryGrau.Add("		  DATAINICIAL")
  QueryGrau.Add("  FROM SAM_PRECOGENERICOGRAU_GRAU")
  QueryGrau.Add(" WHERE HANDLE =:HANDLEGRAU")

  QueryGrau.ParamByName("HANDLEGRAU").AsInteger = HGrau
  QueryGrau.Active = False
  QueryGrau.Active = True

  If Not(QueryGrau.FieldByName("DATAFINAL").IsNull)Then
    MsgBox "Não é possível alterar, vigência já fechada"
    CanContinue = False
    Set INTERFACE = Nothing
    Set QueryGrau = Nothing
    GoTo Fim
  End If
End If

'MsgBox "Handle: "+Str(HGrau)+" - Grau: "+Str(NGrau)

' Jun - SMS 39975 - 09/03/2005 - Inicio
'QueryGrau.Active = False
'QueryGrau.Clear
'QueryGrau.Add("SELECT COUNT(1) QTDE ")
'QueryGrau.Add("  FROM SAM_PRECOGENERICOGRAU_GRAU ")
'QueryGrau.Add(" WHERE GRAU = :HGRAU  ")
'QueryGrau.Add("   AND (PRECOGENERICOGRAU = :HPRECOGENERICO)")

'If CurrentQuery.State = 2 Then
'  QueryGrau.Add("   AND (HANDLE <> :HPRECOGENERICOGRAU) ")
'End If

' Jun - SMS 39975 - 09/03/2005 - Final

'If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
'  QueryGrau.Add("   AND (DATAFINAL IS NULL OR DATAFINAL >= :HDATAINICIAL)  ")
  'Condicao =" AND (DATAFINAL IS NULL OR DATAFINAL >=:TDATAINICIAL)" '"+ CurrentQuery.FieldByName("DATAINICIAL").AsString+"'"
'Else
'  QueryGrau.Add("   AND (DATAFINAL IS NOT NULL)")
  'Condicao =" AND DATAFINAL IS NOT NULL "
'End If

'Andeson sms 26613
'If Not CurrentQuery.FieldByName("EVENTO").IsNull Then
'  QueryGrau.Add("   AND (EVENTO = :HEVENTO OR EVENTO IS NULL) ")
  'Condicao =Condicao +"AND (EVENTO = " +CurrentQuery.FieldByName("EVENTO").AsString +" OR EVENTO IS NULL )"
'End If

'QueryGrau.ParamByName("HGRAU").AsInteger = CurrentQuery.FieldByName("GRAU").AsInteger
'QueryGrau.ParamByName("HPRECOGENERICO").AsInteger = RecordHandleOfTable("SAM_PRECOGENERICOGRAU")

'If CurrentQuery.State = 2 Then
'  QueryGrau.ParamByName("HPRECOGENERICOGRAU").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
'End If

'If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
'  QueryGrau.ParamByName("HDATAINICIAL").AsDateTime = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
'End If

'If Not CurrentQuery.FieldByName("EVENTO").IsNull Then
'  QueryGrau.ParamByName("HEVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
'End If

'QueryGrau.Active = True

'If QueryGrau.FieldByName("QTDE").AsInteger > 0 Then
'  MsgBox("Já existe um registro com vigência aberta")
'  CanContinue = False
'End If

QueryGrau.Active = False
QueryGrau.Clear
QueryGrau.Add("SELECT HANDLE")
QueryGrau.Add("  FROM SAM_PRECOGENERICOGRAU")
QueryGrau.Add(" WHERE HANDLE = " + CStr(RecordHandleOfTable("SAM_PRECOGENERICOGRAU")))
QueryGrau.Active = True

Condicao = " AND PRECOGENERICOGRAU = " + QueryGrau.FieldByName("HANDLE").AsString

If Not CurrentQuery.FieldByName("EVENTO").IsNull Then
  Condicao = Condicao + " AND EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString
End If

'---Alteração Claudemir SMS 18671 -Fim

Linha = INTERFACE.Vigencia(CurrentSystem, "SAM_PRECOGENERICOGRAU_GRAU", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "GRAU", Condicao)

If Linha = "" Then
  CanContinue = True
Else
  CanContinue = False
  MsgBox(Linha)
End If
Set INTERFACE = Nothing

Set QueryGrau = Nothing

Fim :


End Sub




