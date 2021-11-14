# **PROGRAMACION** 


**octubre 11 2021**

trabajamos lo que fueron los ciclo en visual basic de excel "**For**" esto nos perimite hacer unas repeticiones como su nombre lo dice es 
un ciclo que se repite la cantida que pongamos en el codigo seguimos trabajando con ( FOR, TO  ) para trabajar los ciclo 

**EJEMPLO VISUAL BASIC EXCEL**

```
Sub prueba()
    For i = 1 To 10
        MsgBox = "Sena" & i
    Next i
End Sub
```
**EJERCICIO DE VISUAL BASIC EXCEL**

```
Sub ejercicio2()
    a = ("jhon mario")
    For j = 1 To 10 Step 1
        
        Hoja1.Cells(1, 1)(2, 2)(3, 3)(4, 4)(5, 5)(6, 6)(7, 7)(8, 8)(9, 9)(10, 10) = ""
        Hoja1.Cells(11, 11) = a
        Hoja1.Cells(11, 11) = ""
    
    Next j
    
    For m = 1 To 10 Step 1
        Hoja1.Cells(11, 11)(10, 11)(9, 11)(8, 11)(7, 11)(6, 11)(5, 11)(4, 11)(3, 11)(2, 11) = ""
        Hoja1.Cells(1, 11) = a
        Hoja1.Cells(1, 11) = ""
        
    Next m
    
    For m = 1 To 10 Step 1
        Hoja1.Cells(1, 11)(1, 9)(1, 8)(1, 7)(1, 6)(1, 5)(1, 4)(1, 3)(1, 2) = ""
        Hoja1.Cells(1, 1) = a
        
    Next m
End Sub
```