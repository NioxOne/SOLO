'Escrito por Niox - email: nioxdev@gmail.com
'Script para hoja de calculo Calc
'Escribe la hora actual en las celdas de la columna B al ingresar datos en las celdas de la Columna A.
'Util para la entrada de articulos

Sub IHora()
	Dim Doc As Object
	Dim Sheet As Object
	Dim celdaActual As Object
	Dim celdaAN As Object
	Dim celdaBN As Object
	Dim columna As Integer
	Dim fila As Integer
	Dim dataEmpty As Object 
	Dim runMacro As Boolean
	Dim siDato As Boolean	
	
	Doc = ThisComponent
	'Sheet = doc.sheets(0)
	Sheet = Doc.getcurrentcontroller.activesheet
	dataEmpty = com.sun.star.table.CellContentType
	runMacro = True 
	
	Do While runMacro = True
		celdaActual = ThisComponent.CurrentSelection
		On Error Resume Next	
		'Continuar ejecucion si se ha seleccionado un rango de celdas
		columna = celdaActual.cellAddress.Column
		'Obtener posicion de la columna actual  
		If columna = 0 Then 		
		 	'Obtener posicion de la fila actual'
		 	fila = celdaActual.cellAddress.Row
			On Error Resume Next
			'Saber si la fila anterior tiene informacion
			celdaAN = Sheet.getCellByPosition(0,fila-1)
			celdaBN = Sheet.getCellByPosition(1,fila-1)
			siDato = celdaAN.Type <> dataEmpty.EMPTY And CeldaBN.Type = dataEmpty.EMPTY
			If siDato Then
				'Cemtrar el contenido de las celdas'
				celdaAN.ParaAdjust = com.sun.star.style.ParagraphAdjust.CENTER
				tiempo(celdaBN) 'Obtener el tiempo actual'
				celdaBN.ParaAdjust = com.sun.star.style.ParagraphAdjust.CENTER
			End If	
		End If
		wait 25
	Loop
End Sub

'Funcion para obtener el tiempo actual
Function tiempo(ByRef celda As Object)
	celda.FORMULA = time()
	celda.NUMBERFORMAT = 20041
	Celda.VertJustify = com.sun.star.table.CellVertJustify.CENTER
End Function