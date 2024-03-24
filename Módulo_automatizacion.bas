Attribute VB_Name = "Módulo_automatizacion"
'MACRO PARA AUTOMATIZAR LA EXTRACCIÓN DE DATOS DE OTRA FUENTE DE DATOS

Sub AUTOMATIZAR()
    
    'DIMENSIONADO DE VARIABLES
    Dim RUTA As Variant
    Dim ARCHIVO_PRINCIPAL As Variant
    Dim ARCHIVO_FUENTE As Variant
    
    Dim CONTADOR As Integer
    Dim CONTADOR_AUXILIAR As Integer
    
    Dim NOMBRE As String
    Dim FECHA_NACIMIENTO As Date
    Dim CORREO As String
    Dim FECHA_ALTA As Date
    Dim DIRECCION As String
    Dim CP As String
    Dim TELEFONO As String
    Dim GRUPO As String
    
    
    
    'RENOMBRAMOS EL ARCHIVO ORIGINAL
    ARCHIVO_PRINCIPAL = ActiveWorkbook.Name
    
    'ABRIMOS EL EXPLORADOR PARA SELECCIONAR EL ARCHIVO DE ESTUDIO
    RUTA = Application.GetOpenFilename(Title:="SELECCIONA EL ARCHIVO QUE QUIERAS ACTUALIZAR")
    
    'CONTROL DE ERRORES PARA VERIFICAR QUE SE SELECCIONÓ ALGÚN ARCHIVO
    If RUTA <> False Then
    
        'DESACTIVAMOS EL REFRESCO DE PANTALLA PARA AUMENTAR LA PRODUCTIVIDAD DE LA MACRO
        Application.ScreenUpdating = False
        
        'ABRIMOS EL ARCHIVO SELECCIONADO
        Workbooks.Open RUTA
        
        'RENOMBRAMOS EL ARCHIVO DE FUENTE DE DATOS
        ARCHIVO_FUENTE = ActiveWorkbook.Name
        
        'COMENZAMOS EL BUCLE DE BÚSQUEDA DE LAS VARIABLES EN CUESTIÓN
        For CONTADOR = 2 To Workbooks(ARCHIVO_FUENTE).Sheets(1).Cells(Rows.Count, "A").End(xlUp).Row
        
            'ALMACENO LAS VARIABLES QUE QUIERO EXTRAER
            NOMBRE = Workbooks(ARCHIVO_FUENTE).Sheets(1).Cells(CONTADOR, "B").Value
            FECHA_NACIMIENTO = Workbooks(ARCHIVO_FUENTE).Sheets(1).Cells(CONTADOR, "C").Value
            CORREO = Workbooks(ARCHIVO_FUENTE).Sheets(1).Cells(CONTADOR, "G").Value
            FECHA_ALTA = Workbooks(ARCHIVO_FUENTE).Sheets(1).Cells(CONTADOR, "H").Value
            DIRECCION = Workbooks(ARCHIVO_FUENTE).Sheets(1).Cells(CONTADOR, "D").Value
            CP = Workbooks(ARCHIVO_FUENTE).Sheets(1).Cells(CONTADOR, "E").Value
            TELEFONO = Workbooks(ARCHIVO_FUENTE).Sheets(1).Cells(CONTADOR, "F").Value
            
            'FORMATEO LAS VARIABLES DE INTERÉS
                
                'FECHA DE ALTA (SOLAMENTE QUIERO LA FECHA, NO LA HORA)
                FECHA_ALTA = Left(FECHA_ALTA, 10)
                
                'TELÉFONO (QUIERO SOLAMENTE LA ÚLTIMA PARTE)
                CONTADOR_AUXILIAR = InStr(1, TELEFONO, " ", vbTextCompare)
                
                TELEFONO = Right(TELEFONO, Len(TELEFONO) - CONTADOR_AUXILIAR)
                                
            
            'LOS PEGO EN LA FILA CORRESPONDIENTE DE MI HOJA DE CALCULO
            NUEVA_FILA = Workbooks(ARCHIVO_PRINCIPAL).Sheets("REGISTRO").Cells(Rows.Count, "B").End(xlUp).Row + 1
            
            Workbooks(ARCHIVO_PRINCIPAL).Sheets("REGISTRO").Cells(NUEVA_FILA, "B") = NOMBRE
            Workbooks(ARCHIVO_PRINCIPAL).Sheets("REGISTRO").Cells(NUEVA_FILA, "C") = FECHA_NACIMIENTO
            Workbooks(ARCHIVO_PRINCIPAL).Sheets("REGISTRO").Cells(NUEVA_FILA, "D") = CORREO
            Workbooks(ARCHIVO_PRINCIPAL).Sheets("REGISTRO").Cells(NUEVA_FILA, "E") = FECHA_ALTA
            Workbooks(ARCHIVO_PRINCIPAL).Sheets("REGISTRO").Cells(NUEVA_FILA, "F") = DIRECCION
            Workbooks(ARCHIVO_PRINCIPAL).Sheets("REGISTRO").Cells(NUEVA_FILA, "G") = CP
            Workbooks(ARCHIVO_PRINCIPAL).Sheets("REGISTRO").Cells(NUEVA_FILA, "H") = TELEFONO
            
        Next CONTADOR
        
        'CERRAMOS EL ARCHIVO DE FUENTE DE DATOS
        Workbooks(ARCHIVO_FUENTE).Close SAVECHANGES = False
        
    Else
        
        'MENSAJE DE FINALIZACIÓN
        MsgBox "No se seleccionó ningún archivo", vbInformation, "FIN DEL PROGRAMA"
        
    End If
    
End Sub

