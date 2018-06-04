'HASH: 6F6C2EDD1AADDA6386970A4DBC29469B

Option Explicit

Public Sub IMPORTAABRAMGE_OnClick()
  Dim dll As Object
 ' Set dll=CreateBennerObject("SAMUTIL.ROTINAS")
'  dll.Abrange

  Set dll =CreateBennerObject("sampegdigit.digitacao")
  dll.importarAbramge(CurrentSystem,3)'Formato abramge
  Set dll =Nothing
End Sub

Public Sub IMPORTALOTE_OnClick()
  Dim dll As Object
  Set dll =CreateBennerObject("sampegdigit.digitacao")
  dll.importarpeglote(CurrentSystem)
  Set dll =Nothing
End Sub
