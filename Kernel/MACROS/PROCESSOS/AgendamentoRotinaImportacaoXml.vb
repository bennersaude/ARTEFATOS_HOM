'HASH: FA084638254F27B7E2BC8E2C0E4E386C
Option Explicit

Public Sub Main
  Dim componente As CSBusinessComponent
  Set componente = BusinessComponent.CreateInstance("Benner.Saude.Atendimento.Business.SamPericiaRotImportacaoBLL, Benner.Saude.Atendimento.Business")
  componente.Execute("AgendamentoRotinaImportacaoXml")
  Set componente = Nothing
End Sub
