Public KEY_P As String
Public NUM_PRO As String
Public DT_PAG As String
Public TP_DOC As String
Public NUM_DOC_CNPJ As String
Public VL_PAG As Currency
Public COMP_PAG As String
Public NOME_PRO As String
Public NUM_PAR As Inter
Public FOR_PAG As String
Public FOR_PAG_DT As String
Public ISENCAO As String
Public MOTIVO_ISENCAO As String
Public HISTORICO As String
Public lista_par As Variant

'Primeira linha
Pagamento.KEY_P =  range("a" & linha_planilha)
Pagamento.NUM_PRO =  range("b" & linha_planilha)
Pagamento.DT_PAG =  range("c" & linha_planilha)
Pagamento.TP_DOC =  range("d" & linha_planilha)
Pagamento.NUM_DOC_CNPJ = range("e" & linha_planilha)
Pagamento.VL_PAG =  range("f" & linha_planilha)
Pagamento.COMP_PAG =  range("g" & linha_planilha)
'Segunda linha
Pagamento.NOME_PRO =  range("b" & linha_planilha + 1)
Pagamento.NUM_PAR =  range("c" & linha_planilha + 1)
Pagamento.FOR_PAG =  range("d" & linha_planilha + 1)
Pagamento.FOR_PAG_DT =  range("e" & linha_planilha + 1)
'Terceira linha
Pagamento.ISENCAO =  range("b" & linha_planilha +2)
Pagamento.MOTIVO_ISENCAO =  range("c" & linha_planilha+2) 
Pagamento.HISTORICO = range("d" & linha_planilha+2)


Sub PAG_CEPAG_LOCACAO()

Dim linha_final As Integer
Dim linha_planilha As Integer
Dim linha_parcela As Integer
Dim parcelas As Integer
Dim verif_key_p As String
Dim pagamento As Classe1

linha_final = 24 'Mudar

For linha_planilha = 1 To linha_final

    If (linha_planilha = 1 And Range("a" & linha_planilha) <> Null) Then Exit Sub

    If (linha_planilha = 1 Or verif_key_p <> Range("a" & linha_planilha)) Then
        
        Set pagamento = New Classe1
        'Primeira linha
        pagamento.KEY_P = Range("a" & linha_planilha)
        pagamento.NUM_PRO = Range("b" & linha_planilha)
        pagamento.DT_PAG = Range("c" & linha_planilha)
        pagamento.TP_DOC = Range("d" & linha_planilha)
        pagamento.NUM_DOC_CNPJ = Range("e" & linha_planilha)
        pagamento.VL_PAG = Range("f" & linha_planilha)
        pagamento.COMP_PAG = Range("g" & linha_planilha)
        'Segunda linha
        pagamento.NOME_PRO = Range("b" & linha_planilha + 1)
        pagamento.NUM_PAR = Range("c" & linha_planilha + 1)
        pagamento.FOR_PAG = Range("d" & linha_planilha + 1)
        pagamento.FOR_PAG_DT = Range("e" & linha_planilha + 1)
        'Terceira linha
        pagamento.ISENCAO = Range("b" & linha_planilha + 2)
        pagamento.MOTIVO_ISENCAO = Range("c" & linha_planilha + 2)
        pagamento.HISTORICO = Range("d" & linha_planilha + 2)
        
        verif_key_p = pagamento.KEY_P
    
        MsgBox pagamento.KEY_P
    End If

    linha_parcela = linha_planilha + 3

    Do While pagamento.KEY_P = Range("a" & linha_parcela)
        MsgBox Range("b" & linha_parcela) & " ___ " & Range("D" & linha_parcela)
        linha_parcela = linha_parcela + 1
    Loop

    pagamento.lista_par = lista_par
    lista_pag.Add pagamento
    
    linha_planilha = linha_planilha + pagamento.NUM_PAR + 2
    
Next

MsgBox lista_pag.Item(1).lista_par.Item(2).par_vl

End Sub