'HASH: 5A122B0A9D2C2D3C21CD7E0764984CF8
Option Explicit

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("SITUACAOREAJUSTE").AsString = "A" Then
    BOTAOGERAR.Enabled = True
    BOTAOCONFERIR.Enabled = False
    BOTAOENVIARCONFERENCIA.Enabled = False
    BOTAOEFETIVAR.Enabled = False
  ElseIf CurrentQuery.FieldByName("SITUACAOREAJUSTE").AsString = "G" Then
    BOTAOGERAR.Enabled = False
    BOTAOCONFERIR.Enabled = False
    BOTAOENVIARCONFERENCIA.Enabled = True
    BOTAOEFETIVAR.Enabled = False
  ElseIf CurrentQuery.FieldByName("SITUACAOREAJUSTE").AsString = "B" Then
    BOTAOGERAR.Enabled = False
    BOTAOCONFERIR.Enabled = True
    BOTAOENVIARCONFERENCIA.Enabled = False
    BOTAOEFETIVAR.Enabled = False
  ElseIf CurrentQuery.FieldByName("SITUACAOREAJUSTE").AsString = "C" Then
    BOTAOGERAR.Enabled = False
    BOTAOCONFERIR.Enabled = False
    BOTAOENVIARCONFERENCIA.Enabled = False
    BOTAOEFETIVAR.Enabled = True
  ElseIf CurrentQuery.FieldByName("SITUACAOREAJUSTE").AsString = "E" Then
    BOTAOGERAR.Enabled = False
    BOTAOCONFERIR.Enabled = False
    BOTAOENVIARCONFERENCIA.Enabled = False
    BOTAOEFETIVAR.Enabled = False
  End If
End Sub
