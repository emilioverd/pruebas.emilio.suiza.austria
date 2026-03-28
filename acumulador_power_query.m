let
    // ==============================
    // CONFIGURACIÓN
    // ==============================
    CarpetaOrigen = "C:\\Datos\\Excels",
    Extension = ".xlsx",   // También puede ser .xls o .xlsm
    NombreHoja = "Hoja1",   // Cambia por el nombre real de tu hoja

    // ==============================
    // LECTURA DE ARCHIVOS
    // ==============================
    Fuente = Folder.Files(CarpetaOrigen),
    Filtrados = Table.SelectRows(
        Fuente,
        each [Extension] = Extension and ([Attributes]?[Hidden]? <> true)
    ),

    // Cargar el contenido de cada archivo Excel
    ConExcel = Table.AddColumn(
        Filtrados,
        "ExcelData",
        each Excel.Workbook([Content], true),
        type table
    ),

    // Expandir objetos internos (hojas/tablas)
    Expandido = Table.ExpandTableColumn(
        ConExcel,
        "ExcelData",
        {"Name", "Data", "Kind"},
        {"ObjName", "Data", "Kind"}
    ),

    // Quedarse solo con la hoja objetivo
    SoloHoja = Table.SelectRows(
        Expandido,
        each [Kind] = "Sheet" and [ObjName] = NombreHoja
    ),

    // Expandir columnas de datos de forma dinámica
    ColumnasData = if Table.RowCount(SoloHoja) > 0 then Table.ColumnNames(SoloHoja{0}[Data]) else {},
    DatosCombinados = Table.ExpandTableColumn(SoloHoja, "Data", ColumnasData, ColumnasData),

    // Mantener trazabilidad del archivo origen
    Resultado = Table.SelectColumns(
        DatosCombinados,
        List.Combine({{"Name"}, ColumnasData})
    ),
    Renombrado = Table.RenameColumns(Resultado, {{"Name", "ArchivoOrigen"}})
in
    Renombrado
