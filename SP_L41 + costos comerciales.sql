ALTER PROCEDURE SP_L41_Ultimas4Semanas
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @anioActual   INT = YEAR(GETDATE());
    DECLARE @anioAnterior INT = @anioActual - 1;

    DECLARE @tablaActual   NVARCHAR(200) = 'DW.dbo.FacFacturacion_' + CAST(@anioActual AS NVARCHAR(4));
    DECLARE @tablaAnterior NVARCHAR(200) = 'DW.dbo.FacFacturacion_' + CAST(@anioAnterior AS NVARCHAR(4));

    DECLARE @sql NVARCHAR(MAX);

    SET @sql = N'
    ;WITH FactBase AS (
        SELECT
            f.COD_ART,
            f.KGS,
            f.COSTO_ESTIMADO_SEMANAL,
            f.COSTO_ARTICULO_DEVOLUCION_ESTIMADO_SEMANAL,
            f.id_dimcliente,
            f.anulada,
            f.clave_fecha_fact,
            f.fecha_fact,
            dt.CalendarYear,
            dt.WeekNumberOfYear,
            dt.FullDateAlternateKey,
            dt.SpanishMonthName,
            clasif.ID_Clasificacion_Comprobante
        FROM (
            SELECT * FROM ' + @tablaActual + '
            UNION ALL
            SELECT * FROM ' + @tablaAnterior + '
        ) f
        INNER JOIN DW.dbo.DimTime dt
            ON dt.TimeKey = f.clave_fecha_fact
        LEFT JOIN DW.dbo.DimClasificacionComprobante clasif
            ON f.ID_TIPO_COMPROBANTE = clasif.ID_TIPO_COMPROBANTE
        WHERE 
            f.fecha_fact <= CAST(GETDATE() AS date)
            AND (f.anulada IS NULL OR f.anulada = 0)
            AND LEFT(f.id_dimcliente, LEN(f.id_dimcliente)-5) IN (''196'',''336'',''240'')
    ),

    SemanasFact AS (
        SELECT 
            CalendarYear,
            WeekNumberOfYear,
            MIN(FullDateAlternateKey) AS FechaSemana
        FROM FactBase
        GROUP BY CalendarYear, WeekNumberOfYear
    ),

    Semanas4 AS (
        SELECT TOP 4 *
        FROM SemanasFact
        ORDER BY FechaSemana DESC
    ),

    BaseAgregada AS (
        SELECT
            fb.CalendarYear,
            fb.WeekNumberOfYear,
            fb.SpanishMonthName,
            s.FechaSemana,
            fb.COD_ART,
            p.DESCRI_AR,
            SUM(fb.KGS) AS KgsTotales,
            SUM(CASE WHEN fb.ID_Clasificacion_Comprobante = 1 THEN fb.KGS ELSE 0 END) AS KgsClasif1,
            SUM(fb.COSTO_ESTIMADO_SEMANAL)
          + SUM(fb.COSTO_ARTICULO_DEVOLUCION_ESTIMADO_SEMANAL) AS CostoTotal,
            p.UNIDAD_ME,
            p.COD_PRES,
            p.K_BRU_BUL,
            p.CONT_CAJA
        FROM FactBase fb
        INNER JOIN Semanas4 s
            ON fb.CalendarYear = s.CalendarYear
           AND fb.WeekNumberOfYear = s.WeekNumberOfYear
        LEFT JOIN DW.dbo.DimProducto p
            ON p.COD_ART = fb.COD_ART
        WHERE fb.KGS <> 0
        GROUP BY
            fb.CalendarYear,
            fb.WeekNumberOfYear,
            fb.SpanishMonthName,
            s.FechaSemana,
            fb.COD_ART,
            p.DESCRI_AR,
            p.UNIDAD_ME,
            p.COD_PRES,
            p.K_BRU_BUL,
            p.CONT_CAJA
    )

    SELECT
        CalendarYear AS Anio,
        WeekNumberOfYear AS Semana,
        SpanishMonthName AS Mes,
        FechaSemana,
        COD_ART AS Articulo,
        DESCRI_AR,
        KgsTotales,
        KgsClasif1,
        CostoTotal,
        CASE
            WHEN KgsClasif1 = 0 THEN 0
            WHEN UNIDAD_ME = 1 AND COD_PRES = 7 THEN
                (CostoTotal / NULLIF(KgsClasif1,0))
                * (K_BRU_BUL / NULLIF(CONT_CAJA,0))
            WHEN UNIDAD_ME = 1 THEN
                (CostoTotal / NULLIF(KgsClasif1,0))
                * NULLIF(K_BRU_BUL,0)
            ELSE
                (CostoTotal / NULLIF(KgsClasif1,0))
        END AS CostoUnitario
    FROM BaseAgregada
    ORDER BY Anio DESC, Semana DESC, Articulo;
    ';

    -- DEBUG OPCIONAL
    -- SELECT @sql;

    EXEC(@sql);
END;
GO

EXEC DW.dbo.SP_L41_Ultimas4Semanas

SELECT * FROM DATOS..producto WHERE (cod_art = 30)
SELECT * FROM DATOS..v_dw_articost WHERE fecha>'20251101' AND (cod_art = 88)

SELECT * FROM DimProducto
SELECT * FROM DimGranRubro

select * from DW.dbo.facfacturacion where cod_art =30 and  tipo_Comp = 'F' and fecha_fact>'20251101'