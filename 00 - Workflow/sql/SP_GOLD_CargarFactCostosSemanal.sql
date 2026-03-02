USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[SP_GOLD_CargarFactCostosSemanal]    Script Date: 02/03/2026 10:49:41 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER PROCEDURE [dbo].[SP_GOLD_CargarFactCostosSemanal]
AS
BEGIN
    SET NOCOUNT ON;

    TRUNCATE TABLE GOLD.FactCostosSemanal;

    DECLARE @FechaHoy DATE = CAST(GETDATE() AS DATE);

    -- 1️⃣ Base original
    CREATE TABLE #CostosBase (
        Anio INT,
        Semana INT,
        Mes VARCHAR(50),
        UNIDAD_ME INT,
        FechaSemana DATE,
        Articulo INT,
        DESCRI_AR VARCHAR(255),
        GRAN_RUBRO_CDG VARCHAR(50),
        DESCRIPCION_TIPO_ART VARCHAR(100),
        CostoUnitario DECIMAL(18,4),
        CostoComercialUnitario DECIMAL(18,4)
    );

    INSERT INTO #CostosBase
    EXEC dbo.SP_ListarCostos @CantidadSemanas = 260;


    -- 2️⃣ Semana actual calendario
    ;WITH SemanaActual AS (
        SELECT
            MIN(FullDateAlternateKey) AS FechaInicioSemanaActual,
            MAX(FullDateAlternateKey) AS FechaFinSemanaActual
        FROM DW.dbo.DimTime
        WHERE FullDateAlternateKey = @FechaHoy
           OR (
                CalendarYear = (SELECT CalendarYear FROM DW.dbo.DimTime WHERE FullDateAlternateKey = @FechaHoy)
            AND WeekNumberOfYear = (SELECT WeekNumberOfYear FROM DW.dbo.DimTime WHERE FullDateAlternateKey = @FechaHoy)
           )
    ),

    -- 3️⃣ Rango inicial por artículo
    RangoInicialArticulo AS (
        SELECT
            Articulo,
            MIN(FechaSemana) AS FechaMinEvento
        FROM #CostosBase
        GROUP BY Articulo
    ),

    -- 4️⃣ Generación calendario semanal completo
    Semanas AS (
        SELECT
            r.Articulo,
            dt.CalendarYear,
            dt.WeekNumberOfYear,
            MIN(dt.FullDateAlternateKey) AS FechaSemana,     -- lunes
            MAX(dt.FullDateAlternateKey) AS FechaFinSemana   -- domingo
        FROM RangoInicialArticulo r
        INNER JOIN DW.dbo.DimTime dt
            ON dt.FullDateAlternateKey BETWEEN r.FechaMinEvento
                                           AND (SELECT FechaFinSemanaActual FROM SemanaActual)
        GROUP BY
            r.Articulo,
            dt.CalendarYear,
            dt.WeekNumberOfYear
    )

    INSERT INTO GOLD.FactCostosSemanal (
        COD_ART,
        FechaSemana,
        CalendarYear,
        WeekNumberOfYear,
        DESCRI_AR,
        GRAN_RUBRO_CDG,
        UNIDAD_ME,
        CostoUnitario,
        CostoComercialUnitario
    )
    SELECT
        s.Articulo,
        s.FechaSemana,
        s.CalendarYear,
        s.WeekNumberOfYear,
        c.DESCRI_AR,
        c.GRAN_RUBRO_CDG,
        c.UNIDAD_ME,
        c.CostoUnitario,
        c.CostoComercialUnitario
    FROM Semanas s
    OUTER APPLY (
        SELECT TOP 1
            cb.DESCRI_AR,
            cb.GRAN_RUBRO_CDG,
            cb.UNIDAD_ME,
            cb.CostoUnitario,
            cb.CostoComercialUnitario
        FROM #CostosBase cb
        WHERE cb.Articulo = s.Articulo
          AND cb.FechaSemana <= s.FechaFinSemana
          AND cb.CostoUnitario IS NOT NULL
          AND cb.CostoUnitario <> 0
        ORDER BY cb.FechaSemana DESC

    ) c
    WHERE c.CostoUnitario IS NOT NULL;

    DROP TABLE #CostosBase;
END;


