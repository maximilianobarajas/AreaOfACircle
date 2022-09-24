Option Explicit

Sub Macro1()
'
' Macro1 Macro
'

'
End Sub

Sub main()
    'Mi primer programa en VB
    'Programa para calcular el área de un círculo
    
    'Declaración de variables
    Dim area As Double
    Dim radio As Double
    Dim sigue As String
    'Dim area As Double, radio As Double
    'Dim area, radio As Double (No hacer)
    
    'Declaración de constantes
    Const PI = 3.141592653
    
    'Continuidad
    sigue = InputBox("Deseas calcular el área del cículo (S/N)")

    
    'Condición
    While sigue = "S" Or sigue = "s"
        'Inputbox se utiliza para la entrada de datos por teclado
        radio = InputBox("Escribe el valor del radio")
        area = PI * radio * radio
        'Msgbox se utiliza para mostrar datos por pantalla
        MsgBox ("El área del círculo es " & area)
        'Continuidad
        sigue = InputBox("Deseas calcular el área del cículo (S/N)")
    Wend
End Sub