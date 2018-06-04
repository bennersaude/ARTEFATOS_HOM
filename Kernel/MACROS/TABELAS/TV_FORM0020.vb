'HASH: AC7DC56860D60A2AB9BEF318B1FE5080
 


Public Sub TABLE_NewRecord()

Dim qAux1 As Object
Set qAux1 = NewQuery


	qAux1.Add("SELECT B.NIVELDEBUSCA NIVEL1,                               ")
    qAux1.Add("       C.NIVELDEBUSCA NIVEL2,                               ")
    qAux1.Add("       D.NIVELDEBUSCA NIVEL3,                               ")
    qAux1.Add("       E.NIVELDEBUSCA NIVEL4,                               ")
    qAux1.Add("       F.NIVELDEBUSCA NIVEL5,                               ")
    qAux1.Add("       G.NIVELDEBUSCA NIVEL6,                               ")
    qAux1.Add("       H.NIVELDEBUSCA NIVEL7,                               ")
    qAux1.Add("       I.NIVELDEBUSCA NIVEL8,                               ")
    qAux1.Add("       J.NIVELDEBUSCA NIVEL9                                ")
    qAux1.Add("  FROM SAM_CONFIGURABUSCAPRECO A                            ")
    qAux1.Add("  LEFT JOIN SIS_CONFIGURABUSCAPRECO B On (B.Handle = A.NIVEL1)   ")
    qAux1.Add("  LEFT JOIN SIS_CONFIGURABUSCAPRECO C On (C.Handle = A.NIVEL2)   ")
    qAux1.Add("  LEFT JOIN SIS_CONFIGURABUSCAPRECO D On (D.Handle = A.NIVEL3)   ")
    qAux1.Add("  LEFT JOIN SIS_CONFIGURABUSCAPRECO E On (E.Handle = A.NIVEL4)   ")
    qAux1.Add("  LEFT JOIN SIS_CONFIGURABUSCAPRECO F On (F.Handle = A.NIVEL5)   ")
    qAux1.Add("  LEFT JOIN SIS_CONFIGURABUSCAPRECO G On (G.Handle = A.NIVEL6)   ")
    qAux1.Add("  LEFT JOIN SIS_CONFIGURABUSCAPRECO H On (H.Handle = A.NIVEL7)   ")
    qAux1.Add("  LEFT JOIN SIS_CONFIGURABUSCAPRECO I On (I.Handle = A.NIVEL8)   ")
    qAux1.Add("  LEFT JOIN SIS_CONFIGURABUSCAPRECO J On (J.Handle = A.NIVEL9)   ")
	qAux1.Active = True

	'Atribui os Valores
	  CurrentQuery.FieldByName("NIVEL1").AsString = qAux1.FieldByName("NIVEL1").AsString
	  CurrentQuery.FieldByName("NIVEL2").AsString = qAux1.FieldByName("NIVEL2").AsString
      CurrentQuery.FieldByName("NIVEL3").AsString = qAux1.FieldByName("NIVEL3").AsString
      CurrentQuery.FieldByName("NIVEL4").AsString = qAux1.FieldByName("NIVEL4").AsString
      CurrentQuery.FieldByName("NIVEL5").AsString = qAux1.FieldByName("NIVEL5").AsString
      CurrentQuery.FieldByName("NIVEL6").AsString = qAux1.FieldByName("NIVEL6").AsString
      CurrentQuery.FieldByName("NIVEL7").AsString = qAux1.FieldByName("NIVEL7").AsString
      CurrentQuery.FieldByName("NIVEL8").AsString = qAux1.FieldByName("NIVEL8").AsString
      CurrentQuery.FieldByName("NIVEL9").AsString = qAux1.FieldByName("NIVEL9").AsString

	  NIVEL1.ReadOnly = True
	  NIVEL2.ReadOnly = True
	  NIVEL3.ReadOnly = True
	  NIVEL4.ReadOnly = True
	  NIVEL5.ReadOnly = True
	  NIVEL6.ReadOnly = True
	  NIVEL7.ReadOnly = True
	  NIVEL8.ReadOnly = True
	  NIVEL9.ReadOnly = True

End Sub
