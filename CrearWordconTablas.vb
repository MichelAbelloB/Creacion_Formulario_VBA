Sub CrearWordconTablas()
    On Error Resume Next

    Dim objDoc As Object
    Dim objWord As Object
    Set objWord = GetObject(, "Word.Application")
    If objWord Is Nothing Then
        Set objWord = CreateObject("Word.Application")
    End If
    On Error GoTo 0

    objWord.Visible = True

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

    Set objDoc = objWord.Documents.Add

    Set objSection = objDoc.Sections(1)
    Set objHeader = objSection.Headers(1)

    With objHeader.Shapes.AddPicture(Filename:=imgPath, LinkToFile:=False, SaveWithDocument:=True)
        .Left = objWord.CentimetersToPoints(-0.55)
        .Top = objWord.CentimetersToPoints(-0.55)
        .Width = objWord.CentimetersToPoints(3.49)
        .Height = objWord.CentimetersToPoints(1.02)
    End With
    
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    Set rng = ThisWorkbook.Sheets("Formato Unico").Range("B3:B172")
    
    
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    ' Crear tabla
    Dim objTable As Object
    Set objTable = objDoc.Tables.Add(objDoc.Paragraphs(objDoc.Paragraphs.Count).Range, 140, 2)

    ' Configurar el formato de la primera tabla
    With objTable
        .PreferredWidthType = 1
        
        ' Fusionar las celdas de la primera fila
        .Cell(1, 1).Merge .Cell(1, 2)
        .Cell(6, 1).Merge .Cell(6, 2)
        .Cell(14, 1).Merge .Cell(14, 2)
        .Cell(20, 1).Merge .Cell(20, 2)
        .Cell(22, 1).Merge .Cell(22, 2)
        .Cell(52, 1).Merge .Cell(52, 2)
        .Cell(56, 1).Merge .Cell(56, 2)
        .Cell(95, 1).Merge .Cell(95, 2)
        .Cell(121, 1).Merge .Cell(121, 2)
        
        ' Puedes agregar contenido a las celdas de la tabla según sea necesario
        objTable.Cell(1, 1).Range.Text = "FORMATO ÚNICO DE NOTICIA CRIMINAL CONOCIMIENTO INICIAL" & vbCrLf
        objTable.Cell(1, 1).Range.Font.Name = "Cambria"
        objTable.Cell(1, 1).Range.Font.Size = 12
        objTable.Cell(1, 1).Range.ParagraphFormat.Alignment = 1
        
        objTable.Cell(2, 1).Range.Text = "Fecha de Recepción:"
        objTable.Cell(2, 1).Range.Font.Name = "Georgia"
        objTable.Cell(2, 1).Range.Font.Size = 11
        objTable.Cell(2, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(2, 2).Range.Text = rng.Cells(1, 1).Value
        objTable.Cell(2, 2).Range.Font.Name = "Georgia"
        objTable.Cell(2, 2).Range.Font.Size = 11
        objTable.Cell(2, 2).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(3, 1).Range.Text = "Hora:"
        objTable.Cell(3, 1).Range.Font.Name = "Georgia"
        objTable.Cell(3, 1).Range.Font.Size = 11
        objTable.Cell(3, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(3, 2).Range.Text = rng.Cells(2, 1).Value
        objTable.Cell(3, 2).Range.Font.Name = "Georgia"
        objTable.Cell(3, 2).Range.Font.Size = 11
        objTable.Cell(3, 2).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(4, 1).Range.Text = "Departamento:"
        objTable.Cell(4, 1).Range.Font.Name = "Georgia"
        objTable.Cell(4, 1).Range.Font.Size = 11
        objTable.Cell(4, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(4, 2).Range.Text = rng.Cells(3, 1).Value
        objTable.Cell(4, 2).Range.Font.Name = "Georgia"
        objTable.Cell(4, 2).Range.Font.Size = 11
        objTable.Cell(4, 2).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(5, 1).Range.Text = "Municipio:"
        objTable.Cell(5, 1).Range.Font.Name = "Georgia"
        objTable.Cell(5, 1).Range.Font.Size = 11
        objTable.Cell(5, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(5, 2).Range.Text = rng.Cells(4, 1).Value
        objTable.Cell(5, 2).Range.Font.Name = "Georgia"
        objTable.Cell(5, 2).Range.Font.Size = 11
        objTable.Cell(5, 2).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(6, 1).Range.Text = vbCrLf & "NÚMERO ÚNICO DE NOTICIA CRIMINAL" & vbCrLf
        objTable.Cell(6, 1).Range.Font.Name = "Cambria"
        objTable.Cell(6, 1).Range.Font.Size = 12
        objTable.Cell(6, 1).Range.ParagraphFormat.Alignment = 1
        
        objTable.Cell(7, 1).Range.Text = "Caso Noticia:"
        objTable.Cell(7, 1).Range.Font.Name = "Georgia"
        objTable.Cell(7, 1).Range.Font.Size = 11
        objTable.Cell(7, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(7, 2).Range.Text = rng.Cells(7, 1).Value
        objTable.Cell(7, 2).Range.Font.Name = "Georgia"
        objTable.Cell(7, 2).Range.Font.Size = 11
        objTable.Cell(7, 2).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(8, 1).Range.Text = "Departamento:"
        objTable.Cell(8, 1).Range.Font.Name = "Georgia"
        objTable.Cell(8, 1).Range.Font.Size = 11
        objTable.Cell(8, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(8, 2).Range.Text = rng.Cells(8, 1).Value
        objTable.Cell(8, 2).Range.Font.Name = "Georgia"
        objTable.Cell(8, 2).Range.Font.Size = 11
        objTable.Cell(8, 2).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(9, 1).Range.Text = "Municipio:"
        objTable.Cell(9, 1).Range.Font.Name = "Georgia"
        objTable.Cell(9, 1).Range.Font.Size = 11
        objTable.Cell(9, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(9, 2).Range.Text = rng.Cells(9, 1).Value
        objTable.Cell(9, 2).Range.Font.Name = "Georgia"
        objTable.Cell(9, 2).Range.Font.Size = 11
        objTable.Cell(9, 2).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(10, 1).Range.Text = "Entidad Receptora:"
        objTable.Cell(10, 1).Range.Font.Name = "Georgia"
        objTable.Cell(10, 1).Range.Font.Size = 11
        objTable.Cell(10, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(10, 2).Range.Text = rng.Cells(10, 1).Value
        objTable.Cell(10, 2).Range.Font.Name = "Georgia"
        objTable.Cell(10, 2).Range.Font.Size = 11
        objTable.Cell(10, 2).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(11, 1).Range.Text = "Unidad Receptora:"
        objTable.Cell(11, 1).Range.Font.Name = "Georgia"
        objTable.Cell(11, 1).Range.Font.Size = 11
        objTable.Cell(11, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(11, 2).Range.Text = rng.Cells(11, 1).Value
        objTable.Cell(11, 2).Range.Font.Name = "Georgia"
        objTable.Cell(11, 2).Range.Font.Size = 11
        objTable.Cell(11, 2).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(12, 1).Range.Text = "Año:"
        objTable.Cell(12, 1).Range.Font.Name = "Georgia"
        objTable.Cell(12, 1).Range.Font.Size = 11
        objTable.Cell(12, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(12, 2).Range.Text = rng.Cells(12, 1).Value
        objTable.Cell(12, 2).Range.Font.Name = "Georgia"
        objTable.Cell(12, 2).Range.Font.Size = 11
        objTable.Cell(12, 2).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(13, 1).Range.Text = "Consecutivo:"
        objTable.Cell(13, 1).Range.Font.Name = "Georgia"
        objTable.Cell(13, 1).Range.Font.Size = 11
        objTable.Cell(13, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(13, 2).Range.Text = rng.Cells(13, 1).Value
        objTable.Cell(13, 2).Range.Font.Name = "Georgia"
        objTable.Cell(13, 2).Range.Font.Size = 11
        objTable.Cell(13, 2).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(14, 1).Range.Text = vbCrLf & "TIPO DE NOTICIA" & vbCrLf
        objTable.Cell(14, 1).Range.Font.Name = "Cambria"
        objTable.Cell(14, 1).Range.Font.Size = 12
        objTable.Cell(14, 1).Range.ParagraphFormat.Alignment = 1
        
        objTable.Cell(15, 1).Range.Text = "Tipo de Noticia:"
        objTable.Cell(15, 1).Range.Font.Name = "Georgia"
        objTable.Cell(15, 1).Range.Font.Size = 11
        objTable.Cell(15, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(15, 2).Range.Text = rng.Cells(16, 1).Value
        objTable.Cell(15, 2).Range.Font.Name = "Georgia"
        objTable.Cell(15, 2).Range.Font.Size = 11
        objTable.Cell(15, 2).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(16, 1).Range.Text = "Delito Referente:"
        objTable.Cell(16, 1).Range.Font.Name = "Georgia"
        objTable.Cell(16, 1).Range.Font.Size = 11
        objTable.Cell(16, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(16, 2).Range.Text = rng.Cells(17, 1).Value
        objTable.Cell(16, 2).Range.Font.Name = "Georgia"
        objTable.Cell(16, 2).Range.Font.Size = 11
        objTable.Cell(16, 2).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(17, 1).Range.Text = "Modo de operación del delito:"
        objTable.Cell(17, 1).Range.Font.Name = "Georgia"
        objTable.Cell(17, 1).Range.Font.Size = 11
        objTable.Cell(17, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(17, 2).Range.Text = rng.Cells(18, 1).Value
        objTable.Cell(17, 2).Range.Font.Name = "Georgia"
        objTable.Cell(17, 2).Range.Font.Size = 11
        objTable.Cell(17, 2).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(18, 1).Range.Text = "Grado del delito:"
        objTable.Cell(18, 1).Range.Font.Name = "Georgia"
        objTable.Cell(18, 1).Range.Font.Size = 11
        objTable.Cell(18, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(18, 2).Range.Text = rng.Cells(19, 1).Value
        objTable.Cell(18, 2).Range.Font.Name = "Georgia"
        objTable.Cell(18, 2).Range.Font.Size = 11
        objTable.Cell(18, 2).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(19, 1).Range.Text = "Ley de Aplicabilidad:"
        objTable.Cell(19, 1).Range.Font.Name = "Georgia"
        objTable.Cell(19, 1).Range.Font.Size = 11
        objTable.Cell(19, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(19, 2).Range.Text = rng.Cells(20, 1).Value
        objTable.Cell(19, 2).Range.Font.Name = "Georgia"
        objTable.Cell(19, 2).Range.Font.Size = 11
        objTable.Cell(19, 2).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(20, 1).Range.Text = vbCrLf & "AUTORIDADES" & vbCrLf
        objTable.Cell(20, 1).Range.Font.Name = "Cambria"
        objTable.Cell(20, 1).Range.Font.Size = 12
        objTable.Cell(20, 1).Range.ParagraphFormat.Alignment = 1
        
        objTable.Cell(21, 1).Range.Text = "¿El usuario es remitido por una Entidad?:"
        objTable.Cell(21, 1).Range.Font.Name = "Georgia"
        objTable.Cell(21, 1).Range.Font.Size = 11
        objTable.Cell(21, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(21, 2).Range.Text = rng.Cells(23, 1).Value
        objTable.Cell(21, 2).Range.Font.Name = "Georgia"
        objTable.Cell(21, 2).Range.Font.Size = 11
        objTable.Cell(21, 2).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(22, 1).Range.Text = vbCrLf & "DATOS DEL DENUNCIANTE O QUERELLANTE" & vbCrLf
        objTable.Cell(22, 1).Range.Font.Name = "Cambria"
        objTable.Cell(22, 1).Range.Font.Size = 12
        objTable.Cell(22, 1).Range.ParagraphFormat.Alignment = 1
        
        objTable.Cell(23, 1).Range.Text = "Tipo de Documento:"
        objTable.Cell(23, 1).Range.Font.Name = "Georgia"
        objTable.Cell(23, 1).Range.Font.Size = 11
        objTable.Cell(23, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(23, 2).Range.Text = rng.Cells(27, 1).Value
        objTable.Cell(23, 2).Range.Font.Name = "Georgia"
        objTable.Cell(23, 2).Range.Font.Size = 11
        objTable.Cell(23, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(24, 1).Range.Text = "Número de Documento:"
        objTable.Cell(24, 1).Range.Font.Name = "Georgia"
        objTable.Cell(24, 1).Range.Font.Size = 11
        objTable.Cell(24, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(24, 2).Range.Text = rng.Cells(28, 1).Value
        objTable.Cell(24, 2).Range.Font.Name = "Georgia"
        objTable.Cell(24, 2).Range.Font.Size = 11
        objTable.Cell(24, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(25, 1).Range.Text = "Fecha de Expedición:"
        objTable.Cell(25, 1).Range.Font.Name = "Georgia"
        objTable.Cell(25, 1).Range.Font.Size = 11
        objTable.Cell(25, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(25, 2).Range.Text = rng.Cells(29, 1).Value
        objTable.Cell(25, 2).Range.Font.Name = "Georgia"
        objTable.Cell(25, 2).Range.Font.Size = 11
        objTable.Cell(25, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(26, 1).Range.Text = "País de Expedición:"
        objTable.Cell(26, 1).Range.Font.Name = "Georgia"
        objTable.Cell(26, 1).Range.Font.Size = 11
        objTable.Cell(26, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(26, 2).Range.Text = rng.Cells(30, 1).Value
        objTable.Cell(26, 2).Range.Font.Name = "Georgia"
        objTable.Cell(26, 2).Range.Font.Size = 11
        objTable.Cell(26, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(27, 1).Range.Text = "Departamento de Expedición:"
        objTable.Cell(27, 1).Range.Font.Name = "Georgia"
        objTable.Cell(27, 1).Range.Font.Size = 11
        objTable.Cell(27, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(27, 2).Range.Text = rng.Cells(31, 1).Value
        objTable.Cell(27, 2).Range.Font.Name = "Georgia"
        objTable.Cell(27, 2).Range.Font.Size = 11
        objTable.Cell(27, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(28, 1).Range.Text = "Ciudad de Expedición:"
        objTable.Cell(28, 1).Range.Font.Name = "Georgia"
        objTable.Cell(28, 1).Range.Font.Size = 11
        objTable.Cell(28, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(28, 2).Range.Text = rng.Cells(32, 1).Value
        objTable.Cell(28, 2).Range.Font.Name = "Georgia"
        objTable.Cell(28, 2).Range.Font.Size = 11
        objTable.Cell(28, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(29, 1).Range.Text = "Primer Nombre:"
        objTable.Cell(29, 1).Range.Font.Name = "Georgia"
        objTable.Cell(29, 1).Range.Font.Size = 11
        objTable.Cell(29, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(29, 2).Range.Text = rng.Cells(33, 1).Value
        objTable.Cell(29, 2).Range.Font.Name = "Georgia"
        objTable.Cell(29, 2).Range.Font.Size = 11
        objTable.Cell(29, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(30, 1).Range.Text = "Segundo Nombre:"
        objTable.Cell(30, 1).Range.Font.Name = "Georgia"
        objTable.Cell(30, 1).Range.Font.Size = 11
        objTable.Cell(30, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(30, 2).Range.Text = rng.Cells(34, 1).Value
        objTable.Cell(30, 2).Range.Font.Name = "Georgia"
        objTable.Cell(30, 2).Range.Font.Size = 11
        objTable.Cell(30, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(31, 1).Range.Text = "Primer Apellido:"
        objTable.Cell(31, 1).Range.Font.Name = "Georgia"
        objTable.Cell(31, 1).Range.Font.Size = 11
        objTable.Cell(31, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(31, 2).Range.Text = rng.Cells(35, 1).Value
        objTable.Cell(31, 2).Range.Font.Name = "Georgia"
        objTable.Cell(31, 2).Range.Font.Size = 11
        objTable.Cell(31, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(32, 1).Range.Text = "Segundo Apellido:"
        objTable.Cell(32, 1).Range.Font.Name = "Georgia"
        objTable.Cell(32, 1).Range.Font.Size = 11
        objTable.Cell(32, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(32, 2).Range.Text = rng.Cells(36, 1).Value
        objTable.Cell(32, 2).Range.Font.Name = "Georgia"
        objTable.Cell(32, 2).Range.Font.Size = 11
        objTable.Cell(32, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(33, 1).Range.Text = "País de Nacimiento:"
        objTable.Cell(33, 1).Range.Font.Name = "Georgia"
        objTable.Cell(33, 1).Range.Font.Size = 11
        objTable.Cell(33, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(33, 2).Range.Text = rng.Cells(37, 1).Value
        objTable.Cell(33, 2).Range.Font.Name = "Georgia"
        objTable.Cell(33, 2).Range.Font.Size = 11
        objTable.Cell(33, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(34, 1).Range.Text = "Departamento de Nacimiento:"
        objTable.Cell(34, 1).Range.Font.Name = "Georgia"
        objTable.Cell(34, 1).Range.Font.Size = 11
        objTable.Cell(34, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(34, 2).Range.Text = rng.Cells(38, 1).Value
        objTable.Cell(34, 2).Range.Font.Name = "Georgia"
        objTable.Cell(34, 2).Range.Font.Size = 11
        objTable.Cell(34, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(35, 1).Range.Text = "Municipio de Nacimiento:"
        objTable.Cell(35, 1).Range.Font.Name = "Georgia"
        objTable.Cell(35, 1).Range.Font.Size = 11
        objTable.Cell(35, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(35, 2).Range.Text = rng.Cells(39, 1).Value
        objTable.Cell(35, 2).Range.Font.Name = "Georgia"
        objTable.Cell(35, 2).Range.Font.Size = 11
        objTable.Cell(35, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(36, 1).Range.Text = "Fecha de Nacimiento:"
        objTable.Cell(36, 1).Range.Font.Name = "Georgia"
        objTable.Cell(36, 1).Range.Font.Size = 11
        objTable.Cell(36, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(36, 2).Range.Text = rng.Cells(40, 1).Value
        objTable.Cell(36, 2).Range.Font.Name = "Georgia"
        objTable.Cell(36, 2).Range.Font.Size = 11
        objTable.Cell(36, 2).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(37, 1).Range.Text = "Edad:"
        objTable.Cell(37, 1).Range.Font.Name = "Georgia"
        objTable.Cell(37, 1).Range.Font.Size = 11
        objTable.Cell(37, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(37, 2).Range.Text = rng.Cells(41, 1).Value
        objTable.Cell(37, 2).Range.Font.Name = "Georgia"
        objTable.Cell(37, 2).Range.Font.Size = 11
        objTable.Cell(37, 2).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(38, 1).Range.Text = "Sexo:"
        objTable.Cell(38, 1).Range.Font.Name = "Georgia"
        objTable.Cell(38, 1).Range.Font.Size = 11
        objTable.Cell(38, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(38, 2).Range.Text = rng.Cells(42, 1).Value
        objTable.Cell(38, 2).Range.Font.Name = "Georgia"
        objTable.Cell(38, 2).Range.Font.Size = 11
        objTable.Cell(38, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(39, 1).Range.Text = "Tiene alguna discapacidad:"
        objTable.Cell(39, 1).Range.Font.Name = "Georgia"
        objTable.Cell(39, 1).Range.Font.Size = 11
        objTable.Cell(39, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(39, 2).Range.Text = rng.Cells(43, 1).Value
        objTable.Cell(39, 2).Range.Font.Name = "Georgia"
        objTable.Cell(39, 2).Range.Font.Size = 11
        objTable.Cell(39, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(40, 1).Range.Text = "Pertenece a alguna de las poblaciones de especial protección:"
        objTable.Cell(40, 1).Range.Font.Name = "Georgia"
        objTable.Cell(40, 1).Range.Font.Size = 11
        objTable.Cell(40, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(40, 2).Range.Text = rng.Cells(44, 1).Value
        objTable.Cell(40, 2).Range.Font.Name = "Georgia"
        objTable.Cell(40, 2).Range.Font.Size = 11
        objTable.Cell(40, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(41, 1).Range.Text = "Población:"
        objTable.Cell(41, 1).Range.Font.Name = "Georgia"
        objTable.Cell(41, 1).Range.Font.Size = 11
        objTable.Cell(41, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(41, 2).Range.Text = rng.Cells(45, 1).Value
        objTable.Cell(41, 2).Range.Font.Name = "Georgia"
        objTable.Cell(41, 2).Range.Font.Size = 11
        objTable.Cell(41, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(42, 1).Range.Text = "Pueblo o comunidad a la que  pertenece:"
        objTable.Cell(42, 1).Range.Font.Name = "Georgia"
        objTable.Cell(42, 1).Range.Font.Size = 11
        objTable.Cell(42, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(42, 2).Range.Text = rng.Cells(46, 1).Value
        objTable.Cell(42, 2).Range.Font.Name = "Georgia"
        objTable.Cell(42, 2).Range.Font.Size = 11
        objTable.Cell(42, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(43, 1).Range.Text = "Dirección Residencia:"
        objTable.Cell(43, 1).Range.Font.Name = "Georgia"
        objTable.Cell(43, 1).Range.Font.Size = 11
        objTable.Cell(43, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(43, 2).Range.Text = rng.Cells(47, 1).Value
        objTable.Cell(43, 2).Range.Font.Name = "Georgia"
        objTable.Cell(43, 2).Range.Font.Size = 11
        objTable.Cell(43, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(44, 1).Range.Text = "Dirección de Correspondencia:"
        objTable.Cell(44, 1).Range.Font.Name = "Georgia"
        objTable.Cell(44, 1).Range.Font.Size = 11
        objTable.Cell(44, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(44, 2).Range.Text = rng.Cells(48, 1).Value
        objTable.Cell(44, 2).Range.Font.Name = "Georgia"
        objTable.Cell(44, 2).Range.Font.Size = 11
        objTable.Cell(44, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(45, 1).Range.Text = "Complemento Dirección de Correspondencia:"
        objTable.Cell(45, 1).Range.Font.Name = "Georgia"
        objTable.Cell(45, 1).Range.Font.Size = 11
        objTable.Cell(45, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(45, 2).Range.Text = rng.Cells(49, 1).Value
        objTable.Cell(45, 2).Range.Font.Name = "Georgia"
        objTable.Cell(45, 2).Range.Font.Size = 11
        objTable.Cell(45, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(46, 1).Range.Text = "País de Correspondencia:"
        objTable.Cell(46, 1).Range.Font.Name = "Georgia"
        objTable.Cell(46, 1).Range.Font.Size = 11
        objTable.Cell(46, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(46, 2).Range.Text = rng.Cells(50, 1).Value
        objTable.Cell(46, 2).Range.Font.Name = "Georgia"
        objTable.Cell(46, 2).Range.Font.Size = 11
        objTable.Cell(46, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(47, 1).Range.Text = "Departamento de Correspondencia:"
        objTable.Cell(47, 1).Range.Font.Name = "Georgia"
        objTable.Cell(47, 1).Range.Font.Size = 11
        objTable.Cell(47, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(47, 2).Range.Text = rng.Cells(51, 1).Value
        objTable.Cell(47, 2).Range.Font.Name = "Georgia"
        objTable.Cell(47, 2).Range.Font.Size = 11
        objTable.Cell(47, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(48, 1).Range.Text = "Teléfono Celular:"
        objTable.Cell(48, 1).Range.Font.Name = "Georgia"
        objTable.Cell(48, 1).Range.Font.Size = 11
        objTable.Cell(48, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(48, 2).Range.Text = rng.Cells(52, 1).Value
        objTable.Cell(48, 2).Range.Font.Name = "Georgia"
        objTable.Cell(48, 2).Range.Font.Size = 11
        objTable.Cell(48, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(49, 1).Range.Text = "Teléfono Fijo:"
        objTable.Cell(49, 1).Range.Font.Name = "Georgia"
        objTable.Cell(49, 1).Range.Font.Size = 11
        objTable.Cell(49, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(49, 2).Range.Text = rng.Cells(53, 1).Value
        objTable.Cell(49, 2).Range.Font.Name = "Georgia"
        objTable.Cell(49, 2).Range.Font.Size = 11
        objTable.Cell(49, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(50, 1).Range.Text = "Correo Electrónico:"
        objTable.Cell(50, 1).Range.Font.Name = "Georgia"
        objTable.Cell(50, 1).Range.Font.Size = 11
        objTable.Cell(50, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(50, 2).Range.Text = rng.Cells(54, 1).Value
        objTable.Cell(50, 2).Range.Font.Name = "Georgia"
        objTable.Cell(50, 2).Range.Font.Size = 11
        objTable.Cell(50, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(51, 1).Range.Text = "Por qué Medio Desea ser Contactado:"
        objTable.Cell(51, 1).Range.Font.Name = "Georgia"
        objTable.Cell(51, 1).Range.Font.Size = 11
        objTable.Cell(51, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(51, 2).Range.Text = rng.Cells(55, 1).Value
        objTable.Cell(51, 2).Range.Font.Name = "Georgia"
        objTable.Cell(51, 2).Range.Font.Size = 11
        objTable.Cell(51, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(52, 1).Range.Text = vbCrLf & "VÍCTIMAS" & vbCrLf
        objTable.Cell(52, 1).Range.Font.Name = "Cambria"
        objTable.Cell(52, 1).Range.Font.Size = 12
        objTable.Cell(52, 1).Range.ParagraphFormat.Alignment = 1
        
        objTable.Cell(53, 1).Range.Text = "¿Tiene información sobre la(s) victimas(s)?:"
        objTable.Cell(53, 1).Range.Font.Name = "Georgia"
        objTable.Cell(53, 1).Range.Font.Size = 11
        objTable.Cell(53, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(53, 2).Range.Text = rng.Cells(61, 1).Value
        objTable.Cell(53, 2).Range.Font.Name = "Georgia"
        objTable.Cell(53, 2).Range.Font.Size = 11
        objTable.Cell(53, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(54, 1).Range.Text = "¿Cuántas personas fueron víctimas  del delito?:"
        objTable.Cell(54, 1).Range.Font.Name = "Georgia"
        objTable.Cell(54, 1).Range.Font.Size = 11
        objTable.Cell(54, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(54, 2).Range.Text = rng.Cells(62, 1).Value
        objTable.Cell(54, 2).Range.Font.Name = "Georgia"
        objTable.Cell(54, 2).Range.Font.Size = 11
        objTable.Cell(54, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(55, 1).Range.Text = "¿De cuántas de estas víctimas tiene información para aportar?:"
        objTable.Cell(55, 1).Range.Font.Name = "Georgia"
        objTable.Cell(55, 1).Range.Font.Size = 11
        objTable.Cell(55, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(55, 2).Range.Text = rng.Cells(63, 1).Value
        objTable.Cell(55, 2).Range.Font.Name = "Georgia"
        objTable.Cell(55, 2).Range.Font.Size = 11
        objTable.Cell(55, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(56, 1).Range.Text = vbCrLf & "DATOS DE LA VÍCTIMA" & vbCrLf
        objTable.Cell(56, 1).Range.Font.Name = "Cambria"
        objTable.Cell(56, 1).Range.Font.Size = 12
        objTable.Cell(56, 1).Range.ParagraphFormat.Alignment = 1
        
        objTable.Cell(57, 1).Range.Text = "Tipo de Documento:"
        objTable.Cell(57, 1).Range.Font.Name = "Georgia"
        objTable.Cell(57, 1).Range.Font.Size = 11
        objTable.Cell(57, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(57, 2).Range.Text = rng.Cells(67, 1).Value
        objTable.Cell(57, 2).Range.Font.Name = "Georgia"
        objTable.Cell(57, 2).Range.Font.Size = 11
        objTable.Cell(57, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(58, 1).Range.Text = "Número de Documento:"
        objTable.Cell(58, 1).Range.Font.Name = "Georgia"
        objTable.Cell(58, 1).Range.Font.Size = 11
        objTable.Cell(58, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(58, 2).Range.Text = rng.Cells(68, 1).Value
        objTable.Cell(58, 2).Range.Font.Name = "Georgia"
        objTable.Cell(58, 2).Range.Font.Size = 11
        objTable.Cell(58, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(59, 1).Range.Text = "Fecha de Expedición:"
        objTable.Cell(59, 1).Range.Font.Name = "Georgia"
        objTable.Cell(59, 1).Range.Font.Size = 11
        objTable.Cell(59, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(59, 2).Range.Text = rng.Cells(69, 1).Value
        objTable.Cell(59, 2).Range.Font.Name = "Georgia"
        objTable.Cell(59, 2).Range.Font.Size = 11
        objTable.Cell(59, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(60, 1).Range.Text = "País de Expedición:"
        objTable.Cell(60, 1).Range.Font.Name = "Georgia"
        objTable.Cell(60, 1).Range.Font.Size = 11
        objTable.Cell(60, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(60, 2).Range.Text = rng.Cells(70, 1).Value
        objTable.Cell(60, 2).Range.Font.Name = "Georgia"
        objTable.Cell(60, 2).Range.Font.Size = 11
        objTable.Cell(60, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(61, 1).Range.Text = "Departamento de Expedición:"
        objTable.Cell(61, 1).Range.Font.Name = "Georgia"
        objTable.Cell(61, 1).Range.Font.Size = 11
        objTable.Cell(61, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(61, 2).Range.Text = rng.Cells(71, 1).Value
        objTable.Cell(61, 2).Range.Font.Name = "Georgia"
        objTable.Cell(61, 2).Range.Font.Size = 11
        objTable.Cell(61, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(62, 1).Range.Text = "Ciudad de Expedición:"
        objTable.Cell(62, 1).Range.Font.Name = "Georgia"
        objTable.Cell(62, 1).Range.Font.Size = 11
        objTable.Cell(62, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(62, 2).Range.Text = rng.Cells(72, 1).Value
        objTable.Cell(62, 2).Range.Font.Name = "Georgia"
        objTable.Cell(62, 2).Range.Font.Size = 11
        objTable.Cell(62, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(63, 1).Range.Text = "Primer Nombre:"
        objTable.Cell(63, 1).Range.Font.Name = "Georgia"
        objTable.Cell(63, 1).Range.Font.Size = 11
        objTable.Cell(63, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(63, 2).Range.Text = rng.Cells(73, 1).Value
        objTable.Cell(63, 2).Range.Font.Name = "Georgia"
        objTable.Cell(63, 2).Range.Font.Size = 11
        objTable.Cell(63, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(64, 1).Range.Text = "Segundo Nombre:"
        objTable.Cell(64, 1).Range.Font.Name = "Georgia"
        objTable.Cell(64, 1).Range.Font.Size = 11
        objTable.Cell(64, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(64, 2).Range.Text = rng.Cells(74, 1).Value
        objTable.Cell(64, 2).Range.Font.Name = "Georgia"
        objTable.Cell(64, 2).Range.Font.Size = 11
        objTable.Cell(64, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(65, 1).Range.Text = "Primer Apellido:"
        objTable.Cell(65, 1).Range.Font.Name = "Georgia"
        objTable.Cell(65, 1).Range.Font.Size = 11
        objTable.Cell(65, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(65, 2).Range.Text = rng.Cells(75, 1).Value
        objTable.Cell(65, 2).Range.Font.Name = "Georgia"
        objTable.Cell(65, 2).Range.Font.Size = 11
        objTable.Cell(65, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(66, 1).Range.Text = "Segundo Apellido:"
        objTable.Cell(66, 1).Range.Font.Name = "Georgia"
        objTable.Cell(66, 1).Range.Font.Size = 11
        objTable.Cell(66, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(66, 2).Range.Text = rng.Cells(76, 1).Value
        objTable.Cell(66, 2).Range.Font.Name = "Georgia"
        objTable.Cell(66, 2).Range.Font.Size = 11
        objTable.Cell(66, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(67, 1).Range.Text = "País de Nacimiento:"
        objTable.Cell(67, 1).Range.Font.Name = "Georgia"
        objTable.Cell(67, 1).Range.Font.Size = 11
        objTable.Cell(67, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(67, 2).Range.Text = rng.Cells(77, 1).Value
        objTable.Cell(67, 2).Range.Font.Name = "Georgia"
        objTable.Cell(67, 2).Range.Font.Size = 11
        objTable.Cell(67, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(68, 1).Range.Text = "Departamento de Nacimiento:"
        objTable.Cell(68, 1).Range.Font.Name = "Georgia"
        objTable.Cell(68, 1).Range.Font.Size = 11
        objTable.Cell(68, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(68, 2).Range.Text = rng.Cells(78, 1).Value
        objTable.Cell(68, 2).Range.Font.Name = "Georgia"
        objTable.Cell(68, 2).Range.Font.Size = 11
        objTable.Cell(68, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(69, 1).Range.Text = "Municipio de Nacimiento:"
        objTable.Cell(69, 1).Range.Font.Name = "Georgia"
        objTable.Cell(69, 1).Range.Font.Size = 11
        objTable.Cell(69, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(69, 2).Range.Text = rng.Cells(79, 1).Value
        objTable.Cell(69, 2).Range.Font.Name = "Georgia"
        objTable.Cell(69, 2).Range.Font.Size = 11
        objTable.Cell(69, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(70, 1).Range.Text = "Fecha de Nacimiento:"
        objTable.Cell(70, 1).Range.Font.Name = "Georgia"
        objTable.Cell(70, 1).Range.Font.Size = 11
        objTable.Cell(70, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(70, 2).Range.Text = rng.Cells(80, 1).Value
        objTable.Cell(70, 2).Range.Font.Name = "Georgia"
        objTable.Cell(70, 2).Range.Font.Size = 11
        objTable.Cell(70, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(71, 1).Range.Text = "Edad:"
        objTable.Cell(71, 1).Range.Font.Name = "Georgia"
        objTable.Cell(71, 1).Range.Font.Size = 11
        objTable.Cell(71, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(71, 2).Range.Text = rng.Cells(81, 1).Value
        objTable.Cell(71, 2).Range.Font.Name = "Georgia"
        objTable.Cell(71, 2).Range.Font.Size = 11
        objTable.Cell(71, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(72, 1).Range.Text = "Sexo:"
        objTable.Cell(72, 1).Range.Font.Name = "Georgia"
        objTable.Cell(72, 1).Range.Font.Size = 11
        objTable.Cell(72, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(72, 2).Range.Text = rng.Cells(82, 1).Value
        objTable.Cell(72, 2).Range.Font.Name = "Georgia"
        objTable.Cell(72, 2).Range.Font.Size = 11
        objTable.Cell(72, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(73, 1).Range.Text = "Alias:"
        objTable.Cell(73, 1).Range.Font.Name = "Georgia"
        objTable.Cell(73, 1).Range.Font.Size = 11
        objTable.Cell(73, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(73, 2).Range.Text = rng.Cells(83, 1).Value
        objTable.Cell(73, 2).Range.Font.Name = "Georgia"
        objTable.Cell(73, 2).Range.Font.Size = 11
        objTable.Cell(73, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(74, 1).Range.Text = "Tiene alguna discapacidad:"
        objTable.Cell(74, 1).Range.Font.Name = "Georgia"
        objTable.Cell(74, 1).Range.Font.Size = 11
        objTable.Cell(74, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(74, 2).Range.Text = rng.Cells(84, 1).Value
        objTable.Cell(74, 2).Range.Font.Name = "Georgia"
        objTable.Cell(74, 2).Range.Font.Size = 11
        objTable.Cell(74, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(75, 1).Range.Text = "Pertenece a alguna de las poblaciones de especial protección:"
        objTable.Cell(75, 1).Range.Font.Name = "Georgia"
        objTable.Cell(75, 1).Range.Font.Size = 11
        objTable.Cell(75, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(75, 2).Range.Text = rng.Cells(85, 1).Value
        objTable.Cell(75, 2).Range.Font.Name = "Georgia"
        objTable.Cell(75, 2).Range.Font.Size = 11
        objTable.Cell(75, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(76, 1).Range.Text = "¿tiene algún acento en particular?:"
        objTable.Cell(76, 1).Range.Font.Name = "Georgia"
        objTable.Cell(76, 1).Range.Font.Size = 11
        objTable.Cell(76, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(76, 2).Range.Text = rng.Cells(86, 1).Value
        objTable.Cell(76, 2).Range.Font.Name = "Georgia"
        objTable.Cell(76, 2).Range.Font.Size = 11
        objTable.Cell(76, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(77, 1).Range.Text = "¿tiene rasgos o características físicas particulares?:"
        objTable.Cell(77, 1).Range.Font.Name = "Georgia"
        objTable.Cell(77, 1).Range.Font.Size = 11
        objTable.Cell(77, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(77, 2).Range.Text = rng.Cells(87, 1).Value
        objTable.Cell(77, 2).Range.Font.Name = "Georgia"
        objTable.Cell(77, 2).Range.Font.Size = 11
        objTable.Cell(77, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(78, 1).Range.Text = "¿tiene algún tatuaje, aretes, anillos, cadenas, ropa u otros accesorios particulares?:"
        objTable.Cell(78, 1).Range.Font.Name = "Georgia"
        objTable.Cell(78, 1).Range.Font.Size = 11
        objTable.Cell(78, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(78, 2).Range.Text = rng.Cells(88, 1).Value
        objTable.Cell(78, 2).Range.Font.Name = "Georgia"
        objTable.Cell(78, 2).Range.Font.Size = 11
        objTable.Cell(78, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(79, 1).Range.Text = "¿Pertenece o ha pertenecido a algún grupo delincuencial?:"
        objTable.Cell(79, 1).Range.Font.Name = "Georgia"
        objTable.Cell(79, 1).Range.Font.Size = 11
        objTable.Cell(79, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(79, 2).Range.Text = rng.Cells(89, 1).Value
        objTable.Cell(79, 2).Range.Font.Name = "Georgia"
        objTable.Cell(79, 2).Range.Font.Size = 11
        objTable.Cell(79, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(80, 1).Range.Text = "Identidad de género:"
        objTable.Cell(80, 1).Range.Font.Name = "Georgia"
        objTable.Cell(80, 1).Range.Font.Size = 11
        objTable.Cell(80, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(80, 2).Range.Text = rng.Cells(90, 1).Value
        objTable.Cell(80, 2).Range.Font.Name = "Georgia"
        objTable.Cell(80, 2).Range.Font.Size = 11
        objTable.Cell(80, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(81, 1).Range.Text = "Calidad:"
        objTable.Cell(81, 1).Range.Font.Name = "Georgia"
        objTable.Cell(81, 1).Range.Font.Size = 11
        objTable.Cell(81, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(81, 2).Range.Text = rng.Cells(91, 1).Value
        objTable.Cell(81, 2).Range.Font.Name = "Georgia"
        objTable.Cell(81, 2).Range.Font.Size = 11
        objTable.Cell(81, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(82, 1).Range.Text = "Nivel Académico:"
        objTable.Cell(82, 1).Range.Font.Name = "Georgia"
        objTable.Cell(82, 1).Range.Font.Size = 11
        objTable.Cell(82, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(82, 2).Range.Text = rng.Cells(92, 1).Value
        objTable.Cell(82, 2).Range.Font.Name = "Georgia"
        objTable.Cell(82, 2).Range.Font.Size = 11
        objTable.Cell(82, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(83, 1).Range.Text = "Oficio:"
        objTable.Cell(83, 1).Range.Font.Name = "Georgia"
        objTable.Cell(83, 1).Range.Font.Size = 11
        objTable.Cell(83, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(83, 2).Range.Text = rng.Cells(93, 1).Value
        objTable.Cell(83, 2).Range.Font.Name = "Georgia"
        objTable.Cell(83, 2).Range.Font.Size = 11
        objTable.Cell(83, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(84, 1).Range.Text = "Profesión:"
        objTable.Cell(84, 1).Range.Font.Name = "Georgia"
        objTable.Cell(84, 1).Range.Font.Size = 11
        objTable.Cell(84, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(84, 2).Range.Text = rng.Cells(94, 1).Value
        objTable.Cell(84, 2).Range.Font.Name = "Georgia"
        objTable.Cell(84, 2).Range.Font.Size = 11
        objTable.Cell(84, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(85, 1).Range.Text = "Dirección de Correspondencia:"
        objTable.Cell(85, 1).Range.Font.Name = "Georgia"
        objTable.Cell(85, 1).Range.Font.Size = 11
        objTable.Cell(85, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(85, 2).Range.Text = rng.Cells(95, 1).Value
        objTable.Cell(85, 2).Range.Font.Name = "Georgia"
        objTable.Cell(85, 2).Range.Font.Size = 11
        objTable.Cell(85, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(86, 1).Range.Text = "País de Correspondencia:"
        objTable.Cell(86, 1).Range.Font.Name = "Georgia"
        objTable.Cell(86, 1).Range.Font.Size = 11
        objTable.Cell(86, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(86, 2).Range.Text = rng.Cells(96, 1).Value
        objTable.Cell(86, 2).Range.Font.Name = "Georgia"
        objTable.Cell(86, 2).Range.Font.Size = 11
        objTable.Cell(86, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(87, 1).Range.Text = "Departamento de Correspondencia:"
        objTable.Cell(87, 1).Range.Font.Name = "Georgia"
        objTable.Cell(87, 1).Range.Font.Size = 11
        objTable.Cell(87, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(87, 2).Range.Text = rng.Cells(97, 1).Value
        objTable.Cell(87, 2).Range.Font.Name = "Georgia"
        objTable.Cell(87, 2).Range.Font.Size = 11
        objTable.Cell(87, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(88, 1).Range.Text = "Municipio de Correspondencia:"
        objTable.Cell(88, 1).Range.Font.Name = "Georgia"
        objTable.Cell(88, 1).Range.Font.Size = 11
        objTable.Cell(88, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(88, 2).Range.Text = rng.Cells(98, 1).Value
        objTable.Cell(88, 2).Range.Font.Name = "Georgia"
        objTable.Cell(88, 2).Range.Font.Size = 11
        objTable.Cell(88, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(89, 1).Range.Text = "Teléfono Celular:"
        objTable.Cell(89, 1).Range.Font.Name = "Georgia"
        objTable.Cell(89, 1).Range.Font.Size = 11
        objTable.Cell(89, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(89, 2).Range.Text = rng.Cells(99, 1).Value
        objTable.Cell(89, 2).Range.Font.Name = "Georgia"
        objTable.Cell(89, 2).Range.Font.Size = 11
        objTable.Cell(89, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(90, 1).Range.Text = "Teléfono Fijo:"
        objTable.Cell(90, 1).Range.Font.Name = "Georgia"
        objTable.Cell(90, 1).Range.Font.Size = 11
        objTable.Cell(90, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(90, 2).Range.Text = rng.Cells(100, 1).Value
        objTable.Cell(90, 2).Range.Font.Name = "Georgia"
        objTable.Cell(90, 2).Range.Font.Size = 11
        objTable.Cell(90, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(91, 1).Range.Text = "Correo Electrónico:"
        objTable.Cell(91, 1).Range.Font.Name = "Georgia"
        objTable.Cell(91, 1).Range.Font.Size = 11
        objTable.Cell(91, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(91, 2).Range.Text = rng.Cells(101, 1).Value
        objTable.Cell(91, 2).Range.Font.Name = "Georgia"
        objTable.Cell(91, 2).Range.Font.Size = 11
        objTable.Cell(91, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(92, 1).Range.Text = "Conoce el lugar en el que vive la víctima (ciudad, barrio, punto de referencia, etc.):"
        objTable.Cell(92, 1).Range.Font.Name = "Georgia"
        objTable.Cell(92, 1).Range.Font.Size = 11
        objTable.Cell(92, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(92, 2).Range.Text = rng.Cells(102, 1).Value
        objTable.Cell(92, 2).Range.Font.Name = "Georgia"
        objTable.Cell(92, 2).Range.Font.Size = 11
        objTable.Cell(92, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(93, 1).Range.Text = "Otro medio de contacto:"
        objTable.Cell(93, 1).Range.Font.Name = "Georgia"
        objTable.Cell(93, 1).Range.Font.Size = 11
        objTable.Cell(93, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(93, 2).Range.Text = rng.Cells(103, 1).Value
        objTable.Cell(93, 2).Range.Font.Name = "Georgia"
        objTable.Cell(93, 2).Range.Font.Size = 11
        objTable.Cell(93, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(94, 1).Range.Text = "Información adicional:"
        objTable.Cell(94, 1).Range.Font.Name = "Georgia"
        objTable.Cell(94, 1).Range.Font.Size = 11
        objTable.Cell(94, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(94, 2).Range.Text = rng.Cells(104, 1).Value
        objTable.Cell(94, 2).Range.Font.Name = "Georgia"
        objTable.Cell(94, 2).Range.Font.Size = 11
        objTable.Cell(94, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(95, 1).Range.Text = vbCrLf & "DATOS DEL TESTIGO" & vbCrLf
        objTable.Cell(95, 1).Range.Font.Name = "Cambria"
        objTable.Cell(95, 1).Range.Font.Size = 12
        objTable.Cell(95, 1).Range.ParagraphFormat.Alignment = 1
        
        objTable.Cell(96, 1).Range.Text = "Tipo de Documento:"
        objTable.Cell(96, 1).Range.Font.Name = "Georgia"
        objTable.Cell(96, 1).Range.Font.Size = 11
        objTable.Cell(96, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(96, 2).Range.Text = rng.Cells(108, 1).Value
        objTable.Cell(96, 2).Range.Font.Name = "Georgia"
        objTable.Cell(96, 2).Range.Font.Size = 11
        objTable.Cell(96, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(97, 1).Range.Text = "Número de Documento:"
        objTable.Cell(97, 1).Range.Font.Name = "Georgia"
        objTable.Cell(97, 1).Range.Font.Size = 11
        objTable.Cell(97, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(97, 2).Range.Text = rng.Cells(109, 1).Value
        objTable.Cell(97, 2).Range.Font.Name = "Georgia"
        objTable.Cell(97, 2).Range.Font.Size = 11
        objTable.Cell(97, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(98, 1).Range.Text = "Fecha de Expedición:"
        objTable.Cell(98, 1).Range.Font.Name = "Georgia"
        objTable.Cell(98, 1).Range.Font.Size = 11
        objTable.Cell(98, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(98, 2).Range.Text = rng.Cells(110, 1).Value
        objTable.Cell(98, 2).Range.Font.Name = "Georgia"
        objTable.Cell(98, 2).Range.Font.Size = 11
        objTable.Cell(98, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(99, 1).Range.Text = "País de Expedición:"
        objTable.Cell(99, 1).Range.Font.Name = "Georgia"
        objTable.Cell(99, 1).Range.Font.Size = 11
        objTable.Cell(99, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(99, 2).Range.Text = rng.Cells(111, 1).Value
        objTable.Cell(99, 2).Range.Font.Name = "Georgia"
        objTable.Cell(99, 2).Range.Font.Size = 11
        objTable.Cell(99, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(100, 1).Range.Text = "Departamento de Expedición:"
        objTable.Cell(100, 1).Range.Font.Name = "Georgia"
        objTable.Cell(100, 1).Range.Font.Size = 11
        objTable.Cell(100, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(100, 2).Range.Text = rng.Cells(112, 1).Value
        objTable.Cell(100, 2).Range.Font.Name = "Georgia"
        objTable.Cell(100, 2).Range.Font.Size = 11
        objTable.Cell(100, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(101, 1).Range.Text = "Ciudad de Expedición:"
        objTable.Cell(101, 1).Range.Font.Name = "Georgia"
        objTable.Cell(101, 1).Range.Font.Size = 11
        objTable.Cell(101, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(101, 2).Range.Text = rng.Cells(113, 1).Value
        objTable.Cell(101, 2).Range.Font.Name = "Georgia"
        objTable.Cell(101, 2).Range.Font.Size = 11
        objTable.Cell(101, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(102, 1).Range.Text = "Primer Nombre:"
        objTable.Cell(102, 1).Range.Font.Name = "Georgia"
        objTable.Cell(102, 1).Range.Font.Size = 11
        objTable.Cell(102, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(102, 2).Range.Text = rng.Cells(114, 1).Value
        objTable.Cell(102, 2).Range.Font.Name = "Georgia"
        objTable.Cell(102, 2).Range.Font.Size = 11
        objTable.Cell(102, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(103, 1).Range.Text = "Segundo Nombre:"
        objTable.Cell(103, 1).Range.Font.Name = "Georgia"
        objTable.Cell(103, 1).Range.Font.Size = 11
        objTable.Cell(103, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(103, 2).Range.Text = rng.Cells(115, 1).Value
        objTable.Cell(103, 2).Range.Font.Name = "Georgia"
        objTable.Cell(103, 2).Range.Font.Size = 11
        objTable.Cell(103, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(104, 1).Range.Text = "Primer Apellido:"
        objTable.Cell(104, 1).Range.Font.Name = "Georgia"
        objTable.Cell(104, 1).Range.Font.Size = 11
        objTable.Cell(104, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(104, 2).Range.Text = rng.Cells(116, 1).Value
        objTable.Cell(104, 2).Range.Font.Name = "Georgia"
        objTable.Cell(104, 2).Range.Font.Size = 11
        objTable.Cell(104, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(105, 1).Range.Text = "Segundo Apellido:"
        objTable.Cell(105, 1).Range.Font.Name = "Georgia"
        objTable.Cell(105, 1).Range.Font.Size = 11
        objTable.Cell(105, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(105, 2).Range.Text = rng.Cells(117, 1).Value
        objTable.Cell(105, 2).Range.Font.Name = "Georgia"
        objTable.Cell(105, 2).Range.Font.Size = 11
        objTable.Cell(105, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(106, 1).Range.Text = "País de Nacimiento:"
        objTable.Cell(106, 1).Range.Font.Name = "Georgia"
        objTable.Cell(106, 1).Range.Font.Size = 11
        objTable.Cell(106, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(106, 2).Range.Text = rng.Cells(118, 1).Value
        objTable.Cell(106, 2).Range.Font.Name = "Georgia"
        objTable.Cell(106, 2).Range.Font.Size = 11
        objTable.Cell(106, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(107, 1).Range.Text = "Departamento de Nacimiento:"
        objTable.Cell(107, 1).Range.Font.Name = "Georgia"
        objTable.Cell(107, 1).Range.Font.Size = 11
        objTable.Cell(107, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(107, 2).Range.Text = rng.Cells(119, 1).Value
        objTable.Cell(107, 2).Range.Font.Name = "Georgia"
        objTable.Cell(107, 2).Range.Font.Size = 11
        objTable.Cell(107, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(108, 1).Range.Text = "Municipio de Nacimiento:"
        objTable.Cell(108, 1).Range.Font.Name = "Georgia"
        objTable.Cell(108, 1).Range.Font.Size = 11
        objTable.Cell(108, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(108, 2).Range.Text = rng.Cells(120, 1).Value
        objTable.Cell(108, 2).Range.Font.Name = "Georgia"
        objTable.Cell(108, 2).Range.Font.Size = 11
        objTable.Cell(108, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(109, 1).Range.Text = "Edad:"
        objTable.Cell(109, 1).Range.Font.Name = "Georgia"
        objTable.Cell(109, 1).Range.Font.Size = 11
        objTable.Cell(109, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(109, 2).Range.Text = rng.Cells(121, 1).Value
        objTable.Cell(109, 2).Range.Font.Name = "Georgia"
        objTable.Cell(109, 2).Range.Font.Size = 11
        objTable.Cell(109, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(110, 1).Range.Text = "Sexo:"
        objTable.Cell(110, 1).Range.Font.Name = "Georgia"
        objTable.Cell(110, 1).Range.Font.Size = 11
        objTable.Cell(110, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(110, 2).Range.Text = rng.Cells(122, 1).Value
        objTable.Cell(110, 2).Range.Font.Name = "Georgia"
        objTable.Cell(110, 2).Range.Font.Size = 11
        objTable.Cell(110, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(111, 1).Range.Text = "Oficio:"
        objTable.Cell(111, 1).Range.Font.Name = "Georgia"
        objTable.Cell(111, 1).Range.Font.Size = 11
        objTable.Cell(111, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(111, 2).Range.Text = rng.Cells(123, 1).Value
        objTable.Cell(111, 2).Range.Font.Name = "Georgia"
        objTable.Cell(111, 2).Range.Font.Size = 11
        objTable.Cell(111, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(112, 1).Range.Text = "Profesión:"
        objTable.Cell(112, 1).Range.Font.Name = "Georgia"
        objTable.Cell(112, 1).Range.Font.Size = 11
        objTable.Cell(112, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(112, 2).Range.Text = rng.Cells(124, 1).Value
        objTable.Cell(112, 2).Range.Font.Name = "Georgia"
        objTable.Cell(112, 2).Range.Font.Size = 11
        objTable.Cell(112, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(113, 1).Range.Text = "Dirección de Correspondencia:"
        objTable.Cell(113, 1).Range.Font.Name = "Georgia"
        objTable.Cell(113, 1).Range.Font.Size = 11
        objTable.Cell(113, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(113, 2).Range.Text = rng.Cells(125, 1).Value
        objTable.Cell(113, 2).Range.Font.Name = "Georgia"
        objTable.Cell(113, 2).Range.Font.Size = 11
        objTable.Cell(113, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(114, 1).Range.Text = "País de Correspondencia:"
        objTable.Cell(114, 1).Range.Font.Name = "Georgia"
        objTable.Cell(114, 1).Range.Font.Size = 11
        objTable.Cell(114, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(114, 2).Range.Text = rng.Cells(126, 1).Value
        objTable.Cell(114, 2).Range.Font.Name = "Georgia"
        objTable.Cell(114, 2).Range.Font.Size = 11
        objTable.Cell(114, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(115, 1).Range.Text = "Departamento de Correspondencia:"
        objTable.Cell(115, 1).Range.Font.Name = "Georgia"
        objTable.Cell(115, 1).Range.Font.Size = 11
        objTable.Cell(115, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(115, 2).Range.Text = rng.Cells(127, 1).Value
        objTable.Cell(115, 2).Range.Font.Name = "Georgia"
        objTable.Cell(115, 2).Range.Font.Size = 11
        objTable.Cell(115, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(116, 1).Range.Text = "Municipio de Correspondencia:"
        objTable.Cell(116, 1).Range.Font.Name = "Georgia"
        objTable.Cell(116, 1).Range.Font.Size = 11
        objTable.Cell(116, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(116, 2).Range.Text = rng.Cells(128, 1).Value
        objTable.Cell(116, 2).Range.Font.Name = "Georgia"
        objTable.Cell(116, 2).Range.Font.Size = 11
        objTable.Cell(116, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(117, 1).Range.Text = "Teléfono Celular:"
        objTable.Cell(117, 1).Range.Font.Name = "Georgia"
        objTable.Cell(117, 1).Range.Font.Size = 11
        objTable.Cell(117, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(117, 2).Range.Text = rng.Cells(129, 1).Value
        objTable.Cell(117, 2).Range.Font.Name = "Georgia"
        objTable.Cell(117, 2).Range.Font.Size = 11
        objTable.Cell(117, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(118, 1).Range.Text = "Teléfono Fijo:"
        objTable.Cell(118, 1).Range.Font.Name = "Georgia"
        objTable.Cell(118, 1).Range.Font.Size = 11
        objTable.Cell(118, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(118, 2).Range.Text = rng.Cells(130, 1).Value
        objTable.Cell(118, 2).Range.Font.Name = "Georgia"
        objTable.Cell(118, 2).Range.Font.Size = 11
        objTable.Cell(118, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(119, 1).Range.Text = "Otro medio de contacto:"
        objTable.Cell(119, 1).Range.Font.Name = "Georgia"
        objTable.Cell(119, 1).Range.Font.Size = 11
        objTable.Cell(119, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(119, 2).Range.Text = rng.Cells(131, 1).Value
        objTable.Cell(119, 2).Range.Font.Name = "Georgia"
        objTable.Cell(119, 2).Range.Font.Size = 11
        objTable.Cell(119, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(120, 1).Range.Text = "Información adicional:"
        objTable.Cell(120, 1).Range.Font.Name = "Georgia"
        objTable.Cell(120, 1).Range.Font.Size = 11
        objTable.Cell(120, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(120, 2).Range.Text = rng.Cells(132, 1).Value
        objTable.Cell(120, 2).Range.Font.Name = "Georgia"
        objTable.Cell(120, 2).Range.Font.Size = 11
        objTable.Cell(120, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(121, 1).Range.Text = vbCrLf & "DATOS SOBRE LOS HECHOS" & vbCrLf
        objTable.Cell(121, 1).Range.Font.Name = "Cambria"
        objTable.Cell(121, 1).Range.Font.Size = 12
        objTable.Cell(121, 1).Range.ParagraphFormat.Alignment = 1
        
        objTable.Cell(122, 1).Range.Text = "Fecha de comisión de los hechos:"
        objTable.Cell(122, 1).Range.Font.Name = "Georgia"
        objTable.Cell(122, 1).Range.Font.Size = 11
        objTable.Cell(122, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(122, 2).Range.Text = rng.Cells(136, 1).Value
        objTable.Cell(122, 2).Range.Font.Name = "Georgia"
        objTable.Cell(122, 2).Range.Font.Size = 11
        objTable.Cell(122, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(123, 1).Range.Text = "Hora:"
        objTable.Cell(123, 1).Range.Font.Name = "Georgia"
        objTable.Cell(123, 1).Range.Font.Size = 11
        objTable.Cell(123, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(123, 2).Range.Text = rng.Cells(137, 1).Value
        objTable.Cell(123, 2).Range.Font.Name = "Georgia"
        objTable.Cell(123, 2).Range.Font.Size = 11
        objTable.Cell(123, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(124, 1).Range.Text = "Para delitos de acción continuada:"
        objTable.Cell(124, 1).Range.Font.Name = "Georgia"
        objTable.Cell(124, 1).Range.Font.Size = 11
        objTable.Cell(124, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(124, 2).Range.Text = rng.Cells(138, 1).Value
        objTable.Cell(124, 2).Range.Font.Name = "Georgia"
        objTable.Cell(124, 2).Range.Font.Size = 11
        objTable.Cell(124, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(125, 1).Range.Text = "Fecha inicial de comisión:"
        objTable.Cell(125, 1).Range.Font.Name = "Georgia"
        objTable.Cell(125, 1).Range.Font.Size = 11
        objTable.Cell(125, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(125, 2).Range.Text = rng.Cells(139, 1).Value
        objTable.Cell(125, 2).Range.Font.Name = "Georgia"
        objTable.Cell(125, 2).Range.Font.Size = 11
        objTable.Cell(125, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(126, 1).Range.Text = "Hora:"
        objTable.Cell(126, 1).Range.Font.Name = "Georgia"
        objTable.Cell(126, 1).Range.Font.Size = 11
        objTable.Cell(126, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(126, 2).Range.Text = rng.Cells(140, 1).Value
        objTable.Cell(126, 2).Range.Font.Name = "Georgia"
        objTable.Cell(126, 2).Range.Font.Size = 11
        objTable.Cell(126, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(127, 1).Range.Text = "Fecha final de comisión:"
        objTable.Cell(127, 1).Range.Font.Name = "Georgia"
        objTable.Cell(127, 1).Range.Font.Size = 11
        objTable.Cell(127, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(127, 2).Range.Text = rng.Cells(141, 1).Value
        objTable.Cell(127, 2).Range.Font.Name = "Georgia"
        objTable.Cell(127, 2).Range.Font.Size = 11
        objTable.Cell(127, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(128, 1).Range.Text = "Hora:"
        objTable.Cell(128, 1).Range.Font.Name = "Georgia"
        objTable.Cell(128, 1).Range.Font.Size = 11
        objTable.Cell(128, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(128, 2).Range.Text = rng.Cells(142, 1).Value
        objTable.Cell(128, 2).Range.Font.Name = "Georgia"
        objTable.Cell(128, 2).Range.Font.Size = 11
        objTable.Cell(128, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(129, 1).Range.Text = "Lugar de comisión de los hechos:"
        objTable.Cell(129, 1).Range.Font.Name = "Georgia"
        objTable.Cell(129, 1).Range.Font.Size = 11
        objTable.Cell(129, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(129, 2).Range.Text = rng.Cells(143, 1).Value
        objTable.Cell(129, 2).Range.Font.Name = "Georgia"
        objTable.Cell(129, 2).Range.Font.Size = 11
        objTable.Cell(129, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(130, 1).Range.Text = "Departamento:"
        objTable.Cell(130, 1).Range.Font.Name = "Georgia"
        objTable.Cell(130, 1).Range.Font.Size = 11
        objTable.Cell(130, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(130, 2).Range.Text = rng.Cells(144, 1).Value
        objTable.Cell(130, 2).Range.Font.Name = "Georgia"
        objTable.Cell(130, 2).Range.Font.Size = 11
        objTable.Cell(130, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(131, 1).Range.Text = "Municipio:"
        objTable.Cell(131, 1).Range.Font.Name = "Georgia"
        objTable.Cell(131, 1).Range.Font.Size = 11
        objTable.Cell(131, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(131, 2).Range.Text = rng.Cells(145, 1).Value
        objTable.Cell(131, 2).Range.Font.Name = "Georgia"
        objTable.Cell(131, 2).Range.Font.Size = 11
        objTable.Cell(131, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(132, 1).Range.Text = "Localidad o Zona:"
        objTable.Cell(132, 1).Range.Font.Name = "Georgia"
        objTable.Cell(132, 1).Range.Font.Size = 11
        objTable.Cell(132, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(132, 2).Range.Text = rng.Cells(146, 1).Value
        objTable.Cell(132, 2).Range.Font.Name = "Georgia"
        objTable.Cell(132, 2).Range.Font.Size = 11
        objTable.Cell(132, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(133, 1).Range.Text = "Barrio:"
        objTable.Cell(133, 1).Range.Font.Name = "Georgia"
        objTable.Cell(133, 1).Range.Font.Size = 11
        objTable.Cell(133, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(133, 2).Range.Text = rng.Cells(147, 1).Value
        objTable.Cell(133, 2).Range.Font.Name = "Georgia"
        objTable.Cell(133, 2).Range.Font.Size = 11
        objTable.Cell(133, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(134, 1).Range.Text = "Dirección:"
        objTable.Cell(134, 1).Range.Font.Name = "Georgia"
        objTable.Cell(134, 1).Range.Font.Size = 11
        objTable.Cell(134, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(134, 2).Range.Text = rng.Cells(148, 1).Value
        objTable.Cell(134, 2).Range.Font.Name = "Georgia"
        objTable.Cell(134, 2).Range.Font.Size = 11
        objTable.Cell(134, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(135, 1).Range.Text = "Latitud:"
        objTable.Cell(135, 1).Range.Font.Name = "Georgia"
        objTable.Cell(135, 1).Range.Font.Size = 11
        objTable.Cell(135, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(135, 2).Range.Text = rng.Cells(149, 1).Value
        objTable.Cell(135, 2).Range.Font.Name = "Georgia"
        objTable.Cell(135, 2).Range.Font.Size = 11
        objTable.Cell(135, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(136, 1).Range.Text = "longitud:"
        objTable.Cell(136, 1).Range.Font.Name = "Georgia"
        objTable.Cell(136, 1).Range.Font.Size = 11
        objTable.Cell(136, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(136, 2).Range.Text = rng.Cells(150, 1).Value
        objTable.Cell(136, 2).Range.Font.Name = "Georgia"
        objTable.Cell(136, 2).Range.Font.Size = 11
        objTable.Cell(136, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(137, 1).Range.Text = "¿Uso de armas?:"
        objTable.Cell(137, 1).Range.Font.Name = "Georgia"
        objTable.Cell(137, 1).Range.Font.Size = 11
        objTable.Cell(137, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(137, 2).Range.Text = rng.Cells(151, 1).Value
        objTable.Cell(137, 2).Range.Font.Name = "Georgia"
        objTable.Cell(137, 2).Range.Font.Size = 11
        objTable.Cell(137, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(138, 1).Range.Text = "Uso de sustancias tóxicas:"
        objTable.Cell(138, 1).Range.Font.Name = "Georgia"
        objTable.Cell(138, 1).Range.Font.Size = 11
        objTable.Cell(138, 1).Range.ParagraphFormat.Alignment = 0
        
        objTable.Cell(138, 2).Range.Text = rng.Cells(152, 1).Value
        objTable.Cell(138, 2).Range.Font.Name = "Georgia"
        objTable.Cell(138, 2).Range.Font.Size = 11
        objTable.Cell(138, 1).Range.ParagraphFormat.Alignment = 0
        
    End With
    
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



    ' Establecer la ruta y nombre del archivo de destino
    strNombreArchivo = DesktopPath & "FORMATO ÚNICO DE NOTICIA CRIMINAL CONOCIMIENTO INICIAL.docx"

    ' Guardar el documento
    objDoc.SaveAs2 strNombreArchivo
    objDoc.Close
    Set objDoc = Nothing

    ' Cerrar Word si se creó una nueva instancia
    If objWord.Visible = True Then
        objWord.Quit
    End If
    Set objWord = Nothing

    MsgBox "El documento se ha guardado en:" & vbCrLf & vbCrLf & strNombreArchivo, vbInformation, "Documento Guardado"
End Sub

