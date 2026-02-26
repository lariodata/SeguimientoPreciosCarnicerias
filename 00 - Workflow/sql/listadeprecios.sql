USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[SP_GOLD_CargarFactListaPreciosSemanal]    Script Date: 26/02/2026 11:38:37 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [dbo].[SP_GOLD_CargarFactListaPreciosSemanal]
AS
BEGIN
    SET NOCOUNT ON;

    TRUNCATE TABLE GOLD.FactListaPreciosSemanal;

    DECLARE @FechaHoy DATE = CAST(GETDATE() AS DATE);

    ;WITH ListaBase AS (
        SELECT
            lp.id_lista,
            lp.cod_art,
            lp.precio,
            lp.precioxkg,
            lp.unidad_me,
            CAST(lp.fecha AS DATE) AS fecha_evento_precio,
            dt.CalendarYear,
            dt.WeekNumberOfYear,
            dt.FullDateAlternateKey
        FROM DW.dbo.Lista_Precios_Odoo lp
        INNER JOIN DW.dbo.DimTime dt
            ON dt.FullDateAlternateKey = CAST(lp.fecha AS DATE)
    ),

    -- 📅 semana actual calendario
    SemanaActual AS (
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

    Articulos AS (
        SELECT DISTINCT
            id_lista,
            cod_art
        FROM ListaBase
    ),

    RangoInicialArticulo AS (
        SELECT
            id_lista,
            cod_art,
            MIN(FullDateAlternateKey) AS FechaMinEvento
        FROM ListaBase
        GROUP BY id_lista, cod_art
    ),

    Semanas AS (
        SELECT
            r.id_lista,
            r.cod_art,
            dt.CalendarYear,
            dt.WeekNumberOfYear,
            MIN(dt.FullDateAlternateKey) AS FechaSemana,       -- lunes
            MAX(dt.FullDateAlternateKey) AS FechaFinSemana    -- domingo
        FROM RangoInicialArticulo r
        INNER JOIN DW.dbo.DimTime dt
            ON dt.FullDateAlternateKey BETWEEN r.FechaMinEvento
                                           AND (SELECT FechaFinSemanaActual FROM SemanaActual)
        GROUP BY
            r.id_lista,
            r.cod_art,
            dt.CalendarYear,
            dt.WeekNumberOfYear
    )

    INSERT INTO GOLD.FactListaPreciosSemanal (
        id_lista,
        cod_art,
        CalendarYear,
        WeekNumberOfYear,
        FechaSemana,
        precio,
        precioxkg,
        unidad_me,
        fecha_evento_precio
    )
    SELECT
        s.id_lista,
        s.cod_art,
        s.CalendarYear,
        s.WeekNumberOfYear,
        s.FechaSemana,
        p.precio,
        p.precioxkg,
        p.unidad_me,
        p.fecha_evento_precio
    FROM Semanas s
    OUTER APPLY (
        SELECT TOP 1
            lb.precio,
            lb.precioxkg,
            lb.unidad_me,
            lb.fecha_evento_precio
        FROM ListaBase lb
        WHERE lb.id_lista = s.id_lista
          AND lb.cod_art = s.cod_art
          AND lb.fecha_evento_precio <= s.FechaFinSemana
        ORDER BY lb.fecha_evento_precio DESC
    ) p
    WHERE p.precio IS NOT NULL;
END;
