'HASH: 3ADA2B5885282B876010D506B402A0A4
 
Public Sub REIMPRIMIR_OnClick()
   Dim Obj As Object
   Set Obj =CreateBennerObject("SamImpressao.NotaFiscal")
   Obj.Inicializar(CurrentSystem)
   Obj.Reimprimir(CurrentSystem)
   Obj.Finalizar
   Set Obj =Nothing
End Sub
