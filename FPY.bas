Sub Actualizar1_1()
'
' Actualizar1_1 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Año 6")).Select
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Año4").PivotTables(1). _
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
    ActiveSheet.Shapes.Range(Array("Año 5")).Select
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Año3").PivotTables(1). _
        PivotCache.Refresh
    ActiveSheet.Shapes.Range(Array("Mes 6")).Select
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Mes4").PivotTables(1). _
        PivotCache.Refresh
    ActiveSheet.Shapes.Range(Array("Dia 6")).Select
    ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Dia4").PivotTables(1). _
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
        
    Sheets("FPY").Select
End Sub
