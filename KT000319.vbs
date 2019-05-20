'*** Seccion de declaracion de variables ***
Dim m_vCodemp, m_sNroLot, m_dKeyValue, m_OrdenNroFor, m_sRoute
Dim m_sPath, m_sNameFile, m_sFecha, m_sHora, m_tipo
'*******************************************
Function On_Initialization_Transaction(sErrorMessage)
  m_sNroLot = transaccion.numerodelote
  m_vCodemp = cstr(oWizard.Company.Name)
  m_tipo = "Transfer"
  m_OrdenNroFor = 0
  fechaActual 'Cargo la m_sFecha con fomato YYYYMMDDHHMM
  dKeyValue 'Variables Globales para dicionario m_dKeyValue
  With oWizard.Steps(1).Table
    m_sPath = .Fields("VIRT_TXTPAT").Value
    m_sNameFile = .Fields("VIRT_TXTNAM").Value
  End With
  m_sRoute = m_tipo &"\@NOW\@NROLOTE\"
End Function
Function On_Finish_Transaction(sErrorMessage)
  Dim oFileURI, oFso

  Set oFso = CreateObject("scripting.filesystemObject")
  Set oFileURI = CreateObject("scripting.filesystemObject")

  If oFileURI.FileExists(m_sPath & m_sNameFile) Then
    numerarPedidos m_sPath & m_sNameFile
  End If

End Function
Function On_Finish_Insert(sErrorMessage)
End Function
Function numerarPedidos(sNameFile)

  Dim sFilestreamOrigen, sFilestreamDestino, oFile
  Dim sLineaActual, sNumeroCompActual, sNumeroDeCompAnterior
  Dim arrLinea
  Dim oFso, oFileDestino
  Dim IContadorPedido, IContadorItem

  Set oFso = CreateObject("scripting.filesystemObject")

  Set oFile = oFso.GetFile(sNameFile)
  IContador = 0
  sNumeroDeCompAnterior = ""

  Set sFilestreamOrigen = oFile.OpenAsTextStream(1,-2)

  rutaAsignada 'Crea las carpetas que no estan en el path

  Do While Not sFilestreamOrigen.AtEndOfStream
  	sLineaActual = CStr(sFilestreamOrigen.ReadLine)  'Lee la linea
  	arrLinea = Split(sLineaActual,";")
    sNumeroCompActual = arrLinea(m_OrdenNroFor)

    if sNumeroDeCompAnterior = "" or sNumeroDeCompAnterior <> sNumeroCompActual then
      if sNumeroDeCompAnterior <> "" then
          sFilestreamDestino.close
      End if
      sPathArchivo =  CrearArchivoNuevo(sNumeroCompActual)
      Set oFileDestino = oFso.GetFile(sPathArchivo)
      Set sFilestreamDestino = oFileDestino.OpenAsTextStream (2,-2)

      IContadorItem = 0
      IContadorPedido = IContadorPedido + 1
      sNumeroDeCompAnterior = sNumeroCompActual 'Guardo el numero de pedido actual para compararlo con el siguiente
    End if

    IContadorItem = IContadorItem + 1
    sLineaActual = Replace (CStr(sLineaActual), "@@@@", IContadorPedido) 'reemplaza @ numero actual
    sLineaActual = Replace (CStr(sLineaActual), "####", IContadorItem) 'reemplaza @ numero actual
    sFilestreamDestino.writeLine (sLineaActual)
  Loop
  sFilestreamDestino.close
  sFilestreamOrigen.close

  oFile.Delete

  Set oFso = Nothing
  Set oFile = Nothing
  Set oFileDestino = Nothing

End Function
Function CrearArchivoNuevo(sNumero)
  Dim oFso
  Dim sFileDestino, sNewFile
  m_sPath = m_sPath & m_sRoute
  Set oFso = CreateObject("scripting.filesystemObject")
  sNewFile =  m_tipo&"_"&sNumero&"_"& m_sFecha &"_"&m_sHora&"_"& m_sNroLot&".txt"

  If oFso.FileExists(m_sPath & sNewFile) Then
      oFso.DeleteFile m_sPath & sNewFile
  End If

  Set sFileDestino = oFso.CreateTextFile (m_sPath & sNewFile, true)
  CrearArchivoNuevo = m_sPath & sNewFile
End Function
Sub fechaActual()
  sFechaActual =  Right("0000"+CStr(Year(date)),4)
  sFechaActual = sFechaActual +Right("00"+CStr(Month(date)),2)
  sFechaActual = sFechaActual +Right("00"+CStr(Day(date)),2)
  sHora = Right("00"+CStr(hour(time)),2) +Right("00"+CStr(Minute(time)),2)
  m_sHora = sHora
  m_sFecha = sFechaActual
End Sub
Sub rutaAsignada()
  For each dkey in  m_dKeyValue.keys
    m_sRoute = Replace(m_sRoute,dkey ,m_dKeyValue.item(dkey))
  Next
  Do While InStr(m_sRoute,"\")>0
  	sFolder = Mid(m_sRoute,1,InStr(m_sRoute,"\")-1)
  	m_sRoute = Mid(m_sRoute,InStr(m_sRoute,"\")+1,Len(m_sRoute))
    CrearCarpetas(sFolder)
  Loop
End Sub
Sub CrearCarpetas(sNombre)
  Dim oFso
  Dim sFileDestino

  Set oFso = CreateObject("scripting.filesystemObject")

  If Not oFso.FolderExists(m_sPath & sNombre) Then
    Set sFileDestino = oFso.CreateFolder (m_sPath & sNombre)
  End If
  m_sPath = m_sPath & sNombre

End Sub
Sub dKeyValue()

  Set m_dKeyValue = CreateObject("Scripting.Dictionary")
  m_dKeyValue.add "@NROLOTE", m_sNroLot
  m_dKeyValue.add "@NOW" , m_sFecha
  m_dKeyValue.add "@CODEMP" , m_vCodemp

End Sub
