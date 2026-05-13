USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[SP_ListarCostosComerciales]    Script Date: 13/05/2026 11:01:51 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [dbo].[SP_ListarCostosComerciales]
    @CantidadSemanas INT
AS
BEGIN
SET NOCOUNT ON;

IF @CantidadSemanas IS NULL OR @CantidadSemanas <= 0
    SET @CantidadSemanas = 1;

-------------------------------------------------------
-- ÚLTIMAS SEMANAS CON FACTURACIÓN
-------------------------------------------------------

;WITH UltimasSemanas AS (

SELECT DISTINCT TOP (@CantidadSemanas)
    dt.CalendarYear,
    dt.WeekNumberOfYear
FROM DW.dbo.FacFacturacion f
INNER JOIN DW.dbo.DimTime dt
    ON dt.TimeKey = f.clave_fecha_fact
INNER JOIN DW.dbo.DimCliente dc
    ON f.id_dimcliente = dc.ID
WHERE (f.anulada IS NULL OR f.anulada = 0)
  AND dc.NROREP IN (196, 336, 240, 762, 763)
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
    SUM(
        CASE 
            WHEN clasif.ID_Clasificacion_Comprobante = 2
            THEN COALESCE(f.IMP_SINIVA,0)
            ELSE 0
        END
    ) AS NcFinancieras,
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
INNER JOIN DW.dbo.DimCliente dc
    ON f.id_dimcliente = dc.ID
LEFT JOIN DW.dbo.DimClasificacionComprobante clasif
    ON f.ID_TIPO_COMPROBANTE = clasif.ID_TIPO_COMPROBANTE

WHERE (f.anulada IS NULL OR f.anulada = 0)
  AND dc.NROREP IN (196, 336, 240, 762, 763)

GROUP BY
    dt.CalendarYear,
    dt.WeekNumberOfYear,
    f.COD_ART
)

-------------------------------------------------------
-- RESULTADO FINAL (CON CONVERSIÓN DE UNIDADES)
-------------------------------------------------------

SELECT
    ca.CalendarYear,
    ca.WeekNumberOfYear,
    ca.COD_ART,
    ca.ManoObraDespacho,
    ca.CostoFlete,
    ca.CostoAcarreos,
    ca.ComisionVentas,
    ca.ComisionCobranzas,
    ca.DescuentoComisiones,
    ca.CostoRepositoras,
    ca.AcuerdosFijos,
    ca.Fortalecimiento,
    ca.Impuestos,
    ca.NcFinancieras,

    ---------------------------------------------------
    -- COSTO COMERCIAL TOTAL
    ---------------------------------------------------

    (
        ca.ManoObraDespacho
      + ca.CostoFlete
      + ca.CostoAcarreos
      + ca.ComisionVentas
      + ca.ComisionCobranzas
      + ca.DescuentoComisiones
      + ca.CostoRepositoras
      + ca.AcuerdosFijos
      + ca.Fortalecimiento
      + ca.Impuestos
      - ca.NcFinancieras
    ) AS CostoComercial,

    ca.KgsClasif1,

    ---------------------------------------------------
    -- COSTO COMERCIAL UNITARIO KG
    -- Usa fcnConversionUM para calcular factor de conversión
    ---------------------------------------------------

    CASE
        WHEN ca.KgsClasif1 = 0 THEN NULL

        WHEN p.UNIDAD_ME = 0 THEN
        (
            ca.ManoObraDespacho
          + ca.CostoFlete
          + ca.CostoAcarreos
          + ca.ComisionVentas
          + ca.ComisionCobranzas
          + ca.DescuentoComisiones
          + ca.CostoRepositoras
          + ca.AcuerdosFijos
          + ca.Fortalecimiento
          + ca.Impuestos
          - ca.NcFinancieras
        ) / NULLIF(ca.KgsClasif1, 0)

        WHEN p.UNIDAD_ME = 1 AND p.TIPO_ART = 'D' THEN
        (
            ca.ManoObraDespacho
          + ca.CostoFlete
          + ca.CostoAcarreos
          + ca.ComisionVentas
          + ca.ComisionCobranzas
          + ca.DescuentoComisiones
          + ca.CostoRepositoras
          + ca.AcuerdosFijos
          + ca.Fortalecimiento
          + ca.Impuestos
          - ca.NcFinancieras
        ) / NULLIF(ca.KgsClasif1, 0)

        WHEN p.UNIDAD_ME = 1 THEN
        (
            ca.ManoObraDespacho
          + ca.CostoFlete
          + ca.CostoAcarreos
          + ca.ComisionVentas
          + ca.ComisionCobranzas
          + ca.DescuentoComisiones
          + ca.CostoRepositoras
          + ca.AcuerdosFijos
          + ca.Fortalecimiento
          + ca.Impuestos
          - ca.NcFinancieras
        ) / NULLIF(ca.KgsClasif1, 0) * ISNULL(conv.resultado, 1)

        ELSE
        (
            ca.ManoObraDespacho
          + ca.CostoFlete
          + ca.CostoAcarreos
          + ca.ComisionVentas
          + ca.ComisionCobranzas
          + ca.DescuentoComisiones
          + ca.CostoRepositoras
          + ca.AcuerdosFijos
          + ca.Fortalecimiento
          + ca.Impuestos
          - ca.NcFinancieras
        ) / NULLIF(ca.KgsClasif1, 0)
    END AS CostoComercialUnitario

FROM ComercialAgg ca

LEFT JOIN DW.dbo.DimProducto p
    ON p.COD_ART = ca.COD_ART

-------------------------------------------------------
-- CONVERSIÓN DE UNIDADES (UND → KG)
-------------------------------------------------------

OUTER APPLY Comunes.dbo.fcnConversionUM(
    ca.COD_ART,
    1,
    'UND',
    'KG'
) conv

ORDER BY
    ca.CalendarYear DESC,
    ca.WeekNumberOfYear DESC,
    ca.COD_ART;

END