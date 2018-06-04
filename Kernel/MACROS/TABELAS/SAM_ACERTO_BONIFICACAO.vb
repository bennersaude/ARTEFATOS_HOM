'HASH: 5D32388A01C0AEC0FC6DD94D8D6BA5BC
'Macro: SAM_ACERTO_BONIFICACAO
Option Explicit

'#Uses "*ProcuraBeneficiarioAtivo"
Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraBeneficiarioAtivo(False,ServerDate, Beneficiario.Text)
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("BENEFICIARIO").Value=vHandle
  End If
End Sub

Public Sub TABLE_AfterScroll()
  If (CurrentQuery.FieldByName("ORIGEMACERTO").AsString = "A") _
     Or _
     (Not CurrentQuery.FieldByName("fatura").IsNull) Then
    BENEFICIARIO.ReadOnly = True
    DATAACERTO.ReadOnly = True
    VALORSERVICO.ReadOnly = True
    VALORPF.ReadOnly = True
    VALORCONTRIBUICAO.ReadOnly = True
    DESCRICAO.ReadOnly = True   
  Else
    BENEFICIARIO.ReadOnly = False
    DATAACERTO.ReadOnly = False
    VALORSERVICO.ReadOnly = False
    VALORPF.ReadOnly = False
    VALORCONTRIBUICAO.ReadOnly = False
    DESCRICAO.ReadOnly = False
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  PERMITIDO(CanContinue)
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  PERMITIDO(CanContinue)
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  ORIGEMACERTO.ReadOnly = True
End Sub

Public Sub PERMITIDO(CanContinue As Boolean)
  If (CurrentQuery.FieldByName("ORIGEMACERTO").AsString = "A") _
     Or _
     (Not CurrentQuery.FieldByName("fatura").IsNull) Then
    MsgBox("Operação não permitida")
    CanContinue = False
  End If
End Sub
