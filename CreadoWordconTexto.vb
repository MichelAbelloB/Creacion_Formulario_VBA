Sub WordCreado()
    Dim objWord As Object
    Dim objDoc As Object
    Dim objSection As Object
    Dim objHeader As Object
    Dim rng As Range
    Dim strTexto As String
    Dim strNombreArchivo As String
    Dim imgPath As String

    imgPath = "A:\Unidad\Imagen.png"

    Dim DesktopPath As String
    DesktopPath = Environ("USERPROFILE") & "\Desktop\"

    Set rng = ThisWorkbook.Sheets("Formato Unico").Range("B3:B172")

    Set objWord = CreateObject("Word.Application")
    objWord.Visible = True

    Set objDoc = objWord.Documents.Add

    Set objSection = objDoc.Sections(1)
    Set objHeader = objSection.Headers(1)

    With objHeader.Shapes.AddPicture(Filename:=imgPath, LinkToFile:=False, SaveWithDocument:=True)
        .Left = objWord.CentimetersToPoints(-0.55)
        .Top = objWord.CentimetersToPoints(-0.55)
        .Width = objWord.CentimetersToPoints(3.49)
        .Height = objWord.CentimetersToPoints(1.02)
    End With

    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    With objDoc.Content.Paragraphs.Add
        .Range.Text = vbCrLf & "FORMATO ÚNICO DE NOTICIA CRIMINAL CONOCIMIENTO INICIAL"
        .Range.Font.Name = "Cambria"
        .Range.Font.Size = 12
        .Range.Font.Bold = False
        .Alignment = 0
    End With


    With objDoc.Content.Paragraphs.Add
        .Range.Font.Name = "Georgia"
        .Range.Font.Size = 11
        .Alignment = 1 
        .Range.Font.Bold = True
    End With

    strTexto = vbCrLf & "Fecha de Recepción: " & rng.Cells(1, 1).Value & vbCrLf & _
               "Hora: " & rng.Cells(2, 1).Value & vbCrLf & _
               "Departamento: " & rng.Cells(3, 1).Value & vbCrLf & _
               "Municipio: " & rng.Cells(4, 1).Value & vbCrLf & vbCrLf

    With objDoc.Content.Paragraphs.Add
        .Range.Text = strTexto
        .Format.SpaceAfter = 6
    End With
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    With objDoc.Content.Paragraphs.Add
        .Range.Text = "NÚMERO ÚNICO DE NOTICIA CRIMINAL"
        .Range.Font.Name = "Cambria"
        .Range.Font.Size = 12
        .Range.Font.Bold = False
        .Alignment = 0
    End With

    With objDoc.Content.Paragraphs.Add
        .Range.Font.Name = "Georgia"
        .Range.Font.Size = 11
        .Alignment = 1 
        .Range.Font.Bold = True
    End With

    strTexto = vbCrLf & "Caso Noticia: " & rng.Cells(7, 1).Value & vbCrLf & _
               "Departamento: " & rng.Cells(8, 1).Value & vbCrLf & _
               "Municipio: " & rng.Cells(9, 1).Value & vbCrLf & _
               "Entidad Receptora: " & rng.Cells(10, 1).Value & vbCrLf & _
               "Unidad Receptora: " & rng.Cells(11, 1).Value & vbCrLf & _
               "Año: " & rng.Cells(12, 1).Value & vbCrLf & _
               "Consecutivo: " & rng.Cells(13, 1).Value & vbCrLf

    With objDoc.Content.Paragraphs.Add
        .Range.Text = strTexto
        .Format.SpaceAfter = 6
    End With
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  
    With objDoc.Content.Paragraphs.Add
        .Range.Text = vbCrLf & "TIPO DE NOTICIA"
        .Range.Font.Name = "Cambria"
        .Range.Font.Size = 12
        .Range.Font.Bold = False
        .Alignment = 0 
    End With

    With objDoc.Content.Paragraphs.Add
        .Range.Font.Name = "Georgia"
        .Range.Font.Size = 11
        .Alignment = 1 '
        .Range.Font.Bold = True
    End With

    strTexto = vbCrLf & "Tipo de Noticia: " & rng.Cells(16, 1).Value & vbCrLf & _
               "Delito Referente: " & rng.Cells(17, 1).Value & vbCrLf & _
               "Modo de operación del delito: " & rng.Cells(18, 1).Value & vbCrLf & _
               "Grado del delito: " & rng.Cells(19, 1).Value & vbCrLf & _
               "Ley de Aplicabilidad: " & rng.Cells(20, 1).Value & vbCrLf

    With objDoc.Content.Paragraphs.Add
        .Range.Text = strTexto
        .Format.SpaceAfter = 6
    End With
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    With objDoc.Content.Paragraphs.Add
        .Range.Text = vbCrLf & "AUTORIDADES"
        .Range.Font.Name = "Cambria"
        .Range.Font.Size = 12
        .Range.Font.Bold = False
        .Alignment = 0 
    End With


    With objDoc.Content.Paragraphs.Add
        .Range.Font.Name = "Georgia"
        .Range.Font.Size = 11
        .Alignment = 1 '
        .Range.Font.Bold = True
    End With

    strTexto = vbCrLf & "¿El usuario es remitido por una Entidad?: " & rng.Cells(23, 1).Value & vbCrLf & vbCrLf

    With objDoc.Content.Paragraphs.Add
        .Range.Text = strTexto
        .Format.SpaceAfter = 6
    End With
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    With objDoc.Content.Paragraphs.Add
        .Range.Text = "DATOS DEL DENUNCIANTE O QUERELLANTE"
        .Range.Font.Name = "Cambria"
        .Range.Font.Size = 12
        .Range.Font.Bold = False
        .Alignment = 0 
    End With

    With objDoc.Content.Paragraphs.Add
        .Range.Font.Name = "Georgia"
        .Range.Font.Size = 11
        .Alignment = 1
        .Range.Font.Bold = True
    End With

    strTexto = vbCrLf & "Tipo de Documento: " & rng.Cells(27, 1).Value & vbCrLf & _
               "Número de Documento: " & rng.Cells(28, 1).Value & vbCrLf & _
               "Fecha de Expedición: " & rng.Cells(29, 1).Value & vbCrLf & _
               "País de Expedición: " & rng.Cells(30, 1).Value & vbCrLf & _
               "Departamento de Expedición: " & rng.Cells(31, 1).Value & vbCrLf & _
               "Ciudad de Expedición: " & rng.Cells(32, 1).Value & vbCrLf & _
               "Primer Nombre: " & rng.Cells(33, 1).Value & vbCrLf & _
               "Segundo Nombre: " & rng.Cells(34, 1).Value & vbCrLf & _
               "Primer Apellido: " & rng.Cells(35, 1).Value & vbCrLf & _
               "Segundo Apellido: " & rng.Cells(36, 1).Value & vbCrLf & _
               "País de Nacimiento: " & rng.Cells(37, 1).Value & vbCrLf & _
               "Departamento de Nacimiento: " & rng.Cells(38, 1).Value & vbCrLf & _
               "Municipio de Nacimiento: " & rng.Cells(39, 1).Value & vbCrLf & _
               "Fecha de Nacimiento: " & rng.Cells(40, 1).Value & vbCrLf & _
               "Edad: " & rng.Cells(41, 1).Value & vbCrLf & _
               "Sexo: " & rng.Cells(42, 1).Value & vbCrLf & _
               "Tiene alguna discapacidad: " & rng.Cells(43, 1).Value & vbCrLf & _
               "Pertenece a alguna de las poblaciones de especial protección: " & rng.Cells(44, 1).Value & vbCrLf & _
               "Población: " & rng.Cells(45, 1).Value & vbCrLf & _
               "Pueblo o comunidad a la que pertenece: " & rng.Cells(46, 1).Value & vbCrLf & _
               "Tipo de Dirección: " & rng.Cells(47, 1).Value & vbCrLf & _
               "Dirección de Correspondencia: " & rng.Cells(48, 1).Value & vbCrLf & _
               "Complemento Dirección de Correspondencia: " & rng.Cells(49, 1).Value & vbCrLf & _
               "País de Correspondencia Departamento de Correspondencia: " & rng.Cells(50, 1).Value & vbCrLf & _
               "Municipio de Correspondencia: " & rng.Cells(51, 1).Value & vbCrLf

    strTexto = vbCrLf & "Teléfono Celular: " & rng.Cells(52, 1).Value & vbCrLf & _
               "Teléfono Fijo: " & rng.Cells(53, 1).Value & vbCrLf & _
               "Correo Electrónico: " & rng.Cells(54, 1).Value & vbCrLf & _
               "Por qué Medio Desea ser Contactado: " & rng.Cells(55, 1).Value & vbCrLf & _
               "Estimación de los daños y perjuicios: " & rng.Cells(56, 1).Value & vbCrLf
    
    With objDoc.Content.Paragraphs.Add
        .Range.Text = strTexto
        .Format.SpaceAfter = 6
    End With
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    With objDoc.Content.Paragraphs.Add
        .Range.Text = vbCrLf & "VÍCTIMAS"
        .Range.Font.Name = "Cambria"
        .Range.Font.Size = 12
        .Range.Font.Bold = False
        .Alignment = 0 
        .Format.SpaceAfter = 0
    End With

    With objDoc.Content.Paragraphs.Add
        .Range.Font.Name = "Georgia"
        .Range.Font.Size = 11
        .Alignment = 1
        .Range.Font.Bold = True
    End With

    strTexto = vbCrLf & "¿Tiene información sobre la(s) victimas(s)?: " & rng.Cells(61, 1).Value & vbCrLf & _
               "¿Cuántas personas fueron víctimas del delito?: " & rng.Cells(62, 1).Value & vbCrLf & _
               "¿De cuántas de estas víctimas tiene información para aportar?: " & rng.Cells(63, 1).Value & vbCrLf

    With objDoc.Content.Paragraphs.Add
        .Range.Text = strTexto
        .Format.SpaceAfter = 6
    End With
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    With objDoc.Content.Paragraphs.Add
        .Range.Text = vbCrLf & "DATOS DE LA VÍCTIMA"
        .Range.Font.Name = "Cambria"
        .Range.Font.Size = 12
        .Range.Font.Bold = False
        .Alignment = 0 
        .Format.SpaceAfter = 12 
    End With

    With objDoc.Content.Paragraphs.Add
        .Range.Font.Name = "Georgia"
        .Range.Font.Size = 11
        .Alignment = 1 
        .Range.Font.Bold = True
    End With

    strTexto = vbCrLf & "Tipo de Documento: " & rng.Cells(67, 1).Value & vbCrLf & _
               "Número de Documento: " & rng.Cells(68, 1).Value & vbCrLf & _
               "Fecha de Expedición: " & rng.Cells(69, 1).Value & vbCrLf & _
               "País de Expedición: " & rng.Cells(70, 1).Value & vbCrLf & _
               "Departamento de Expedición: " & rng.Cells(71, 1).Value & vbCrLf & _
               "Ciudad de Expedición: " & rng.Cells(72, 1).Value & vbCrLf & _
               "Primer Nombre: " & rng.Cells(73, 1).Value & vbCrLf & _
               "Segundo Nombre: " & rng.Cells(74, 1).Value & vbCrLf & _
               "Primer Apellido: " & rng.Cells(75, 1).Value & vbCrLf & _
               "Segundo Apellido: " & rng.Cells(76, 1).Value & vbCrLf & _
               "País de Nacimiento: " & rng.Cells(77, 1).Value & vbCrLf & _
               "Departamento de Nacimiento: " & rng.Cells(78, 1).Value & vbCrLf & _
               "Municipio de Nacimiento: " & rng.Cells(79, 1).Value & vbCrLf & _
               "Fecha de Nacimiento: " & rng.Cells(80, 1).Value & vbCrLf & _
               "Edad: " & rng.Cells(81, 1).Value & vbCrLf & _
               "Sexo: " & rng.Cells(82, 1).Value & vbCrLf & _
               "Alias: " & rng.Cells(83, 1).Value & vbCrLf & _
               "Tiene alguna discapacidad: " & rng.Cells(84, 1).Value & vbCrLf & _
               "Pertenece a alguna de las poblaciones de especial protección: " & rng.Cells(85, 1).Value & vbCrLf & _
               "¿tiene algún acento en particular?:  " & rng.Cells(86, 1).Value & vbCrLf & _
               "¿tiene rasgos o características físicas particulares?: " & rng.Cells(87, 1).Value & vbCrLf & _
               "¿tiene algún tatuaje, aretes, anillos, cadenas, ropa u otros accesorios particulares?: " & rng.Cells(88, 1).Value & vbCrLf & _
               "¿Pertenece o ha pertenecido a algún grupo delincuencial?: " & rng.Cells(89, 1).Value & vbCrLf & _
               "Identidad de género: " & rng.Cells(90, 1).Value & vbCrLf & _
               "Calidad: " & rng.Cells(91, 1).Value & vbCrLf
    
    strTexto = "Nivel Académico: " & rng.Cells(92, 1).Value & vbCrLf & _
               "Oficio: " & rng.Cells(93, 1).Value & vbCrLf & _
               "Profesión: " & rng.Cells(94, 1).Value & vbCrLf & _
               "Dirección de Correspondencia: " & rng.Cells(95, 1).Value & vbCrLf & _
               "País de Correspondencia: " & rng.Cells(96, 1).Value & vbCrLf & _
               "Departamento de Correspondencia: " & rng.Cells(97, 1).Value & vbCrLf & _
               "Municipio de Correspondencia: " & rng.Cells(98, 1).Value & vbCrLf & _
               "Teléfono Celular: " & rng.Cells(99, 1).Value & vbCrLf & _
               "Teléfono Fijo: " & rng.Cells(100, 1).Value & vbCrLf & _
               "Correo Electrónico: " & rng.Cells(101, 1).Value & vbCrLf & _
               "Conoce el lugar en el que vive la víctima (ciudad, barrio, punto de referencia, etc.): " & rng.Cells(102, 1).Value & vbCrLf & _
               "Otro medio de contacto: " & rng.Cells(103, 1).Value & vbCrLf & _
               "Información adicional: " & rng.Cells(104, 1).Value & vbCrLf
               
    With objDoc.Content.Paragraphs.Add
        .Range.Text = strTexto
        .Format.SpaceAfter = 6 ' Espacio después del párrafo
    End With
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    With objDoc.Content.Paragraphs.Add
        .Range.Text = vbCrLf & "DATOS DEL TESTIGO"
        .Range.Font.Name = "Cambria"
        .Range.Font.Size = 12
        .Range.Font.Bold = False
        .Alignment = 0 
        .Format.SpaceAfter = 12 
    End With

    With objDoc.Content.Paragraphs.Add
        .Range.Font.Name = "Georgia"
        .Range.Font.Size = 11
        .Alignment = 1 '
        .Range.Font.Bold = True
    End With

    strTexto = vbCrLf & "Tipo de Documento: " & rng.Cells(108, 1).Value & vbCrLf & _
               "Número de Documento: " & rng.Cells(109, 1).Value & vbCrLf & _
               "Fecha de Expedición: " & rng.Cells(110, 1).Value & vbCrLf & _
               "País de Expedición: " & rng.Cells(111, 1).Value & vbCrLf & _
               "Departamento de Expedición: " & rng.Cells(112, 1).Value & vbCrLf & _
               "Ciudad de Expedición: " & rng.Cells(113, 1).Value & vbCrLf & _
               "Primer Nombre: " & rng.Cells(114, 1).Value & vbCrLf & _
               "Segundo Nombre: " & rng.Cells(115, 1).Value & vbCrLf & _
               "Primer Apellido: " & rng.Cells(116, 1).Value & vbCrLf & _
               "Segundo Apellido: " & rng.Cells(117, 1).Value & vbCrLf & _
               "País de Nacimiento: " & rng.Cells(118, 1).Value & vbCrLf & _
               "Departamento de Nacimiento: " & rng.Cells(119, 1).Value & vbCrLf & _
               "Municipio de Nacimiento: " & rng.Cells(120, 1).Value & vbCrLf & _
               "Edad: " & rng.Cells(121, 1).Value & vbCrLf & _
               "Sexo: " & rng.Cells(122, 1).Value & vbCrLf & _
               "Oficio: " & rng.Cells(123, 1).Value & vbCrLf & _
               "Profesión: " & rng.Cells(124, 1).Value & vbCrLf & _
               "Dirección de Correspondencia: " & rng.Cells(125, 1).Value & vbCrLf & _
               "País de Correspondencia: " & rng.Cells(126, 1).Value & vbCrLf & _
               "Departamento de Correspondencia: " & rng.Cells(127, 1).Value & vbCrLf & _
               "Municipio de Correspondencia: " & rng.Cells(128, 1).Value & vbCrLf & _
               "Teléfono Celular: " & rng.Cells(129, 1).Value & vbCrLf & _
               "Teléfono Fijo: " & rng.Cells(130, 1).Value & vbCrLf & _
               "Otro medio de contacto: " & rng.Cells(131, 1).Value & vbCrLf & _
               "Información adicional: " & rng.Cells(132, 1).Value & vbCrLf

    With objDoc.Content.Paragraphs.Add
        .Range.Text = strTexto
        .Format.SpaceAfter = 6 
    End With
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    With objDoc.Content.Paragraphs.Add
        .Range.Text = vbCrLf & "DATOS SOBRE LOS HECHOS"
        .Range.Font.Name = "Cambria"
        .Range.Font.Size = 12
        .Range.Font.Bold = False
        .Alignment = 0 
        .Format.SpaceAfter = 12 
    End With

    With objDoc.Content.Paragraphs.Add
        .Range.Font.Name = "Georgia"
        .Range.Font.Size = 11
        .Alignment = 1 '
        .Range.Font.Bold = True
    End With

    strTexto = vbCrLf & "Fecha de comisión de los hechos: " & rng.Cells(136, 1).Value & vbCrLf & _
               "Hora: " & rng.Cells(137, 1).Value & vbCrLf & _
               "Para delitos de acción continuada: " & rng.Cells(138, 1).Value & vbCrLf & _
               "Fecha inicial de comisión: " & rng.Cells(139, 1).Value & vbCrLf & _
               "Hora: " & rng.Cells(140, 1).Value & vbCrLf & _
               "Fecha final de comisión: " & rng.Cells(141, 1).Value & vbCrLf & _
               "Hora: " & rng.Cells(142, 1).Value & vbCrLf & _
               "Lugar de comisión de los hechos: " & rng.Cells(143, 1).Value & vbCrLf & _
               "Departamento: " & rng.Cells(144, 1).Value & vbCrLf & _
               "Municipio: " & rng.Cells(145, 1).Value & vbCrLf & _
               "Localidad o Zona: " & rng.Cells(146, 1).Value & vbCrLf & _
               "Barrio: " & rng.Cells(147, 1).Value & vbCrLf & _
               "Dirección: " & rng.Cells(148, 1).Value & vbCrLf & _
               "Latitud: " & rng.Cells(149, 1).Value & vbCrLf & _
               "longitud: " & rng.Cells(150, 1).Value & vbCrLf & _
               "¿Uso de armas?: " & rng.Cells(151, 1).Value & vbCrLf & _
               "Uso de sustancias tóxicas: " & rng.Cells(152, 1).Value & vbCrLf

    With objDoc.Content.Paragraphs.Add
        .Range.Text = strTexto
        .Format.SpaceAfter = 6 
    End With
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    strNombreArchivo = DesktopPath & "FORMATO ÚNICO DE NOTICIA CRIMINAL CONOCIMIENTO INICIAL.docx"

    objDoc.SaveAs2 strNombreArchivo
    objDoc.Close
    objWord.Quit
    Set objWord = Nothing

    MsgBox "El documento se ha guardado en:" & vbCrLf & vbCrLf & strNombreArchivo, vbInformation, "Documento Guardado"
End Sub
