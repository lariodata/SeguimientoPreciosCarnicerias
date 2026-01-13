USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[SP_ListasPrecios_UltimasSemanas]    Script Date: 13/01/2026 12:43:44 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [dbo].[SP_ListasPrecios_UltimasSemanas]
    @CantidadSemanas INT,
    @IdLista INT
AS
BEGIN
    SET NOCOUNT ON;

    -- Valores por defecto
    IF @CantidadSemanas IS NULL OR @CantidadSemanas <= 0
        SET @CantidadSemanas = 3;

    IF @IdLista IS NULL
        THROW 50001, 'Debe especificarse un IdLista válido.', 1;

    ;WITH ListaBase AS (
        SELECT 
            lp.id_lista,
            lp.lista,
            lp.cod_art AS Articulo,
            lp.producto AS DESCRI_AR,
            cp.GRAN_RUBRO_CDG,
            lp.unidad_me AS UNIDAD_ME,
            lp.precio,
            lp.precioxkg,
            ta.DESCRIPCION_TIPO_ART,
            dt.CalendarYear,
            dt.WeekNumberOfYear,
            dt.DayNumberOfWeek,
            dt.FullDateAlternateKey,
            dt.SpanishMonthName
        FROM DW.dbo.Lista_Precios_Odoo lp
        INNER JOIN DW.dbo.DimTime dt
            ON dt.FullDateAlternateKey = CAST(lp.fecha AS DATE)
        LEFT JOIN DW.dbo.DimProducto p
            ON p.COD_ART = lp.COD_ART
        LEFT JOIN DW.dbo.DimClasificacionProducto cp
            ON cp.COD_ART = lp.COD_ART
        LEFT JOIN DW.dbo.DimTipoArticulo ta
            ON p.TIPO_ART = ta.CODIGO_TIPO_ART
        WHERE lp.id_lista = @IdLista
    ),

    UltimasSemanas AS (
        SELECT TOP (@CantidadSemanas)
            CalendarYear,
            WeekNumberOfYear,
            MIN(FullDateAlternateKey) AS FechaSemana
        FROM ListaBase
        GROUP BY CalendarYear, WeekNumberOfYear
        ORDER BY FechaSemana DESC
    )

    SELECT
        lb.id_lista,
        lb.lista,
        lb.Articulo,
        lb.DESCRI_AR,
        lb.GRAN_RUBRO_CDG,
        lb.UNIDAD_ME,
        lb.precio,
        lb.precioxkg,
        lb.DESCRIPCION_TIPO_ART,
        us.FechaSemana,
        lb.CalendarYear,
        lb.WeekNumberOfYear,
        lb.SpanishMonthName
    FROM ListaBase lb
    INNER JOIN UltimasSemanas us
        ON lb.CalendarYear = us.CalendarYear
       AND lb.WeekNumberOfYear = us.WeekNumberOfYear
    ORDER BY us.FechaSemana DESC;

END;


