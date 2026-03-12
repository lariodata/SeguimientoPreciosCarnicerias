USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[SP_ListarCostosUnitarios]    Script Date: 12/03/2026 7:24:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [dbo].[SP_ListarCostosUnitarios]
    @CantidadSemanas INT
AS
BEGIN
SET NOCOUNT ON;

IF @CantidadSemanas IS NULL OR @CantidadSemanas <= 0
    SET @CantidadSemanas = 1;

-------------------------------------------------------
-- COSTOS POR SEMANA (QUEDARSE CON EL ÚLTIMO DE LA SEMANA)
-------------------------------------------------------

;WITH CostosSemana AS (

SELECT
    ac.*,

    DATEADD(
        DAY,
        -((DATEPART(WEEKDAY, ac.fecha) + @@DATEFIRST - 2) % 7),
        CAST(ac.fecha AS DATE)
    ) AS FechaSemana,

    ROW_NUMBER() OVER(
        PARTITION BY
            ac.cod_art,
            DATEADD(
                DAY,
                -((DATEPART(WEEKDAY, ac.fecha) + @@DATEFIRST - 2) % 7),
                CAST(ac.fecha AS DATE)
            )
        ORDER BY ac.fecha DESC
    ) AS rn

FROM DW.dbo.articost ac
WHERE ac.tipolista = 'S'

),

-------------------------------------------------------
-- ÚLTIMAS SEMANAS
-------------------------------------------------------

UltimosLunes AS (

SELECT TOP (@CantidadSemanas)
    FechaSemana
FROM CostosSemana
WHERE rn = 1
GROUP BY FechaSemana
ORDER BY FechaSemana DESC

)

-------------------------------------------------------
-- RESULTADO
-------------------------------------------------------

SELECT
    dt.CalendarYear,
    dt.WeekNumberOfYear,
    ul.FechaSemana,

    cs.cod_art AS Articulo,
    p.UNIDAD_ME,
    p.DESCRI_AR,
    cp.GRAN_RUBRO_CDG,
    ta.DESCRIPCION_TIPO_ART,

---------------------------------------------------
-- COSTO UNITARIO
---------------------------------------------------

    CASE
        WHEN p.UNIDAD_ME = 0 THEN
            cs.COSTO

        WHEN p.UNIDAD_ME = 1 THEN
            cs.COSTO * ISNULL(conv.resultado,0)

        ELSE
            cs.COSTO
    END AS CostoUnitario

FROM CostosSemana cs

INNER JOIN UltimosLunes ul
    ON cs.FechaSemana = ul.FechaSemana
   AND cs.rn = 1

LEFT JOIN DW.dbo.DimTime dt
    ON dt.FullDateAlternateKey = ul.FechaSemana

LEFT JOIN DW.dbo.DimProducto p
    ON p.COD_ART = cs.cod_art

LEFT JOIN DW.dbo.DimClasificacionProducto cp
    ON cp.COD_ART = cs.cod_art

LEFT JOIN DW.dbo.DimTipoArticulo ta
    ON p.TIPO_ART = ta.CODIGO_TIPO_ART

-------------------------------------------------------
-- CONVERSIÓN A KG
-------------------------------------------------------

OUTER APPLY Comunes.dbo.fcnConversionUM(
    cs.cod_art,
    1,
    'UND',
    'KG'
) conv

ORDER BY
    ul.FechaSemana DESC,
    cs.cod_art;

END
