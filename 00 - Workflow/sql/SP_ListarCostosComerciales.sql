USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[SP_ListarCostosComerciales]    Script Date: 12/03/2026 7:23:14 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [dbo].[SP_ListarCostosComerciales]
    @CantidadSemanas INT
AS
BEGIN
SET NOCOUNT ON;

--Asigno 1 si no viene CantidadSemanas como parámetro
IF @CantidadSemanas IS NULL OR @CantidadSemanas <= 0
    SET @CantidadSemanas = 1;

-------------------------------------------------------
-- ÚLTIMAS SEMANAS CON FACTURACIÓN
-------------------------------------------------------
--Devuelve algo como:
--2026	11
--2026	10
--2026	9
--2026	8 

;WITH UltimasSemanas AS (

SELECT DISTINCT TOP (@CantidadSemanas)
    dt.CalendarYear,
    dt.WeekNumberOfYear
FROM DW.dbo.FacFacturacion f
INNER JOIN DW.dbo.DimTime dt
    ON dt.TimeKey = f.clave_fecha_fact
WHERE (f.anulada IS NULL OR f.anulada = 0)
ORDER BY
    dt.CalendarYear DESC,
    dt.WeekNumberOfYear DESC
),


-------------------------------------------------------
-- AGREGADO COMERCIAL (DESGLOSADO)
-------------------------------------------------------

ComercialAgg AS (

SELECT
    dt.CalendarYear,
    dt.WeekNumberOfYear,
    f.COD_ART,

    ---------------------------------------------------
    -- COMPONENTES COSTO COMERCIAL
    ---------------------------------------------------

    SUM(COALESCE(f.MANO_OBRA_DESPACHO,0))       AS ManoObraDespacho,
    SUM(COALESCE(f.COSTO_FLETE,0))              AS CostoFlete,
    SUM(COALESCE(f.COSTO_ACARREOS,0))           AS CostoAcarreos,
    SUM(COALESCE(f.COSTO_COMISION_VENTA,0))     AS ComisionVentas,
    SUM(COALESCE(f.COSTO_COMISION_COBRANZAS,0)) AS ComisionCobranzas,
    SUM(COALESCE(f.DESCUENTO_COMISIONES,0))     AS DescuentoComisiones,
    SUM(COALESCE(f.COSTO_REPOSITORAS,0))        AS CostoRepositoras,
    SUM(COALESCE(f.ACUERDOS_FIJOS,0))           AS AcuerdosFijos,
    SUM(COALESCE(f.FORTALECIMIENTO,0))          AS Fortalecimiento,
    SUM(COALESCE(f.IMPUESTOS,0))                AS Impuestos,

    ---------------------------------------------------
    -- NC FINANCIERAS
    ---------------------------------------------------

    SUM(
        CASE 
            WHEN clasif.ID_Clasificacion_Comprobante = 2
            THEN COALESCE(f.IMP_SINIVA,0)
            ELSE 0
        END
    ) AS NcFinancieras,

    ---------------------------------------------------
    -- KGS
    ---------------------------------------------------

    SUM(
        CASE
            WHEN clasif.ID_Clasificacion_Comprobante = 1
            THEN COALESCE(f.KGS,0)
            ELSE 0
        END
    ) AS KgsClasif1

FROM DW.dbo.FacFacturacion f

INNER JOIN DW.dbo.DimTime dt
    ON dt.TimeKey = f.clave_fecha_fact

INNER JOIN UltimasSemanas s
    ON s.CalendarYear = dt.CalendarYear
   AND s.WeekNumberOfYear = dt.WeekNumberOfYear

LEFT JOIN DW.dbo.DimClasificacionComprobante clasif
    ON f.ID_TIPO_COMPROBANTE = clasif.ID_TIPO_COMPROBANTE

WHERE (f.anulada IS NULL OR f.anulada = 0)

GROUP BY
    dt.CalendarYear,
    dt.WeekNumberOfYear,
    f.COD_ART
)

-------------------------------------------------------
-- RESULTADO FINAL
-------------------------------------------------------

SELECT
    CalendarYear,
    WeekNumberOfYear,
    COD_ART,
    ManoObraDespacho,
    CostoFlete,
    CostoAcarreos,
    ComisionVentas,
    ComisionCobranzas,
    DescuentoComisiones,
    CostoRepositoras,
    AcuerdosFijos,
    Fortalecimiento,
    Impuestos,
    NcFinancieras,

    ---------------------------------------------------
    -- COSTO COMERCIAL TOTAL
    ---------------------------------------------------

    (
        ManoObraDespacho
      + CostoFlete
      + CostoAcarreos
      + ComisionVentas
      + ComisionCobranzas
      + DescuentoComisiones
      + CostoRepositoras
      + AcuerdosFijos
      + Fortalecimiento
      + Impuestos
      - NcFinancieras
    ) AS CostoComercial,

    KgsClasif1,

    ---------------------------------------------------
    -- COSTO COMERCIAL UNITARIO KG
    ---------------------------------------------------

    CASE
        WHEN KgsClasif1 = 0 THEN NULL
        ELSE
        (
            ManoObraDespacho
          + CostoFlete
          + CostoAcarreos
          + ComisionVentas
          + ComisionCobranzas
          + DescuentoComisiones
          + CostoRepositoras
          + AcuerdosFijos
          + Fortalecimiento
          + Impuestos
          - NcFinancieras
        ) / NULLIF(KgsClasif1,0)
    END AS CostoComercialUnitario

FROM ComercialAgg

ORDER BY
    CalendarYear DESC,
    WeekNumberOfYear DESC,
    COD_ART;

END
