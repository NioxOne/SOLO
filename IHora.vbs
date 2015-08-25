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
	Dim columnaDatos As Integer
	Dim columnaHora As Integer
	Dim posActual As Integer
	Dim enterMove As String
	Dim num As Integer

	'Indica la columna en la cual se va almacenar la entrada de datos
	'*Cambiar si se requiere utilizar otra columna
	'A = 0, B = 1, C = 3 ...
	columnaDatos = 0

	'Indica la columna en la cual se va registrar la hora de entrada
	columnaHora = columnaDatos + 1

	'Cambiar dependiendo hacia donde se mueve la tecla intro
	' R = Right, D = Down
	enterMove = "R"

	Select Case enterMove
		Case "R"
			posActual = columnaHora
			num = 0
		Case "D"
			posActual = columnaDatos
			num = -1
	End Select
	
	Doc = ThisComponent
	Sheet = Doc.getcurrentcontroller.activesheet
	dataEmpty = com.sun.star.table.CellContentType
	runMacro = True 

	Do While runMacro = True
		celdaActual = ThisComponent.CurrentSelection
		On Error Resume Next	
		'Continuar ejecucion si se ha seleccionado un rango de celdas
		'Obtener posicion de la columna actual 
		columna = celdaActual.cellAddress.Column 
		If columna = posActual Then 		'Cambiar valor de columna y fila 
		 	'Obtener posicion de la fila actual'
		 	fila = celdaActual.cellAddress.Row +  num
			On Error Resume Next
			'Saber si la columna anterior tiene información
			celdaAN = Sheet.getCellByPosition(columnaDatos,fila)
			celdaBN = Sheet.getCellByPosition(columnaHora,fila)
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
End Function
