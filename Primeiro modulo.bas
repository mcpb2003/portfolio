Attribute VB_Name = "Módulo1"
Sub FiltrarEAlterar()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim filtroRange As Range
    Dim dataInicial As Date
    Dim dataFinal As Date
    Dim data2024Inicio As Date
    Dim data2024Fim As Date

    ' Definir a folha ativa (altera conforme necessário)
    Set ws = ActiveSheet

    ' Descobrir a última linha da tabela
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Definir intervalo total
    Set rng = ws.Range("A1").CurrentRegion

    ' Definir datas corretamente
    data2024Inicio = DateValue("01/01/2024")
    data2024Fim = DateValue("31/12/2024")
    dataInicial = DateValue("05/01/2025")
    dataFinal = DateValue("04/02/2025")

    ' Ativar Autofilter
    ws.AutoFilterMode = False

    ' Aplicar filtro na coluna O (F_EFECTO_INICIAL) para mostrar apenas valores de 2024
    rng.AutoFilter Field:=15, _
                   Criteria1:=">=" & CLng(data2024Inicio), _
                   Operator:=xlAnd, _
                   Criteria2:="<=" & CLng(data2024Fim)

    ' Aplicar filtro na coluna X (FECHALTA) para mostrar valores entre 05/01/2025 e 04/02/2025
    rng.AutoFilter Field:=24, _
                   Criteria1:=">=" & CLng(dataInicial), _
                   Operator:=xlAnd, _
                   Criteria2:="<=" & CLng(dataFinal)

    ' Definir intervalo filtrado (excluindo cabeçalhos)
    On Error Resume Next
    Set filtroRange = ws.Range("O2:O" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    ' Alterar todas as células da coluna O (F_EFECTO_INICIAL) no filtro para 01/01/2025
    If Not filtroRange Is Nothing Then
        Dim cel As Range
        For Each cel In filtroRange
            cel.Value = DateValue("01/01/2025")
        Next cel
    End If

    ' Remover filtros
    ws.AutoFilterMode = False

    MsgBox "Processo concluído! Todas as células filtradas foram alteradas para 01/01/2025.", vbInformation, "Finalizado"
End Sub


