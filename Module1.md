Private Sub CommandButton1_Click()
'заполняет матрицу 30х30 рандомными целочисленными значениями'
For i = 1 To 30
For j = 1 To 30
Cells(j, i) = Int((100 * Rnd) - 50)
Next j
Next i
End Sub

Private Sub CommandButton2_Click()
For i = 1 To 30
For j = 1 To 30
If j = i Then 'берутся во внимание элементы главной диагонали матрицы'
a = a + Cells(j, i) 'суммируются значения элементов главной диагонали'
k = k + 1 'считается количество элементов главной диагонали'
End If
b = a / k 'считается среднее арифметическое по формуле'
Next j
Next i
For i = 1 To 5 'берется минор матрицы'
For j = 1 To 5
If Cells(j, i) > b Then 'считается количество миноров матрицы больше ее среднего арифметического'
ans = ans + 1
End If
Next j
Next i
MsgBox (ans)
End Sub

Private Sub CommandButton3_Click()
UserForm1.Hide
End Sub


Private Sub UserForm_Click()

End Sub