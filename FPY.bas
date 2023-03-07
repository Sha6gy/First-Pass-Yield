Attribute VB_Name = "Módulo1"
Sub Actualizar()
Attribute Actualizar.VB_Description = "Actualizacion de entradas para el Yield"
Attribute Actualizar.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Actualizar Macro
' Actualizacion de entradas para el Yield
'

'
    ActiveSheet.Shapes.Range(Array("Año 4")).Select
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Año3").PivotTables(1). _
        PivotCache.Refresh
    ActiveSheet.Shapes.Range(Array("Mes 4")).Select
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Mes3").PivotTables(1). _
        PivotCache.Refresh
    ActiveSheet.Shapes.Range(Array("Dia 4")).Select
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Dia3").PivotTables(1). _
        PivotCache.Refresh
    ActiveSheet.Shapes.Range(Array("Año 3")).Select
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Año2").PivotTables(1). _
        PivotCache.Refresh
    ActiveSheet.Shapes.Range(Array("Mes 3")).Select
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Mes2").PivotTables(1). _
        PivotCache.Refresh
    ActiveSheet.Shapes.Range(Array("Dia 3")).Select
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Dia2").PivotTables(1). _
        PivotCache.Refresh
    Sheets("Charts OP & Equipment").Select
    ActiveSheet.Shapes.Range(Array("Año 1")).Select
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Año").PivotTables(1). _
        PivotCache.Refresh
    ActiveSheet.Shapes.Range(Array("Mes 1")).Select
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Mes").PivotTables(1). _
        PivotCache.Refresh
    ActiveSheet.Shapes.Range(Array("Dia 1")).Select
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Dia").PivotTables(1). _
        PivotCache.Refresh
    ActiveSheet.Shapes.Range(Array("Año 2")).Select
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Año1").PivotTables(1). _
        PivotCache.Refresh
    ActiveSheet.Shapes.Range(Array("Mes 2")).Select
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Mes1").PivotTables(1). _
        PivotCache.Refresh
    ActiveSheet.Shapes.Range(Array("Dia 2")).Select
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Dia1").PivotTables(1). _
        PivotCache.Refresh
    With ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Dia")
        .SlicerItems("5").Selected = True
        .SlicerItems("1").Selected = False
        .SlicerItems("2").Selected = False
        .SlicerItems("3").Selected = False
        .SlicerItems("4").Selected = False
        .SlicerItems("6").Selected = False
        .SlicerItems("7").Selected = False
        .SlicerItems("8").Selected = False
        .SlicerItems("9").Selected = False
        .SlicerItems("10").Selected = False
        .SlicerItems("11").Selected = False
        .SlicerItems("12").Selected = False
        .SlicerItems("13").Selected = False
        .SlicerItems("14").Selected = False
        .SlicerItems("15").Selected = False
        .SlicerItems("16").Selected = False
        .SlicerItems("17").Selected = False
        .SlicerItems("18").Selected = False
        .SlicerItems("19").Selected = False
        .SlicerItems("20").Selected = False
        .SlicerItems("21").Selected = False
        .SlicerItems("22").Selected = False
        .SlicerItems("23").Selected = False
        .SlicerItems("24").Selected = False
        .SlicerItems("25").Selected = False
        .SlicerItems("26").Selected = False
        .SlicerItems("27").Selected = False
        .SlicerItems("28").Selected = False
        .SlicerItems("29").Selected = False
        .SlicerItems("30").Selected = False
        .SlicerItems("31").Selected = False
        .SlicerItems("(en blanco)").Selected = False
    End With
    Sheets("Pivot").Select
    Range("A1").Select
    ActiveSheet.PivotTables("TablaDinámica4").PivotCache.Refresh
    Range("O1").Select
    ActiveSheet.PivotTables("TablaDinámica5").PivotCache.Refresh
    Range("Y3").Select
    ActiveSheet.PivotTables("TablaDinámica6").PivotCache.Refresh
    Range("Y19").Select
    ActiveSheet.PivotTables("TablaDinámica7").PivotCache.Refresh
    
    Sheets("FPY").Select
    
End Sub
