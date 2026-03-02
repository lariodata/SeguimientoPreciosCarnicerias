USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[SP_ListarCostos]    Script Date: 02/03/2026 8:36:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [dbo].[SP_ListarCostos]
    @CantidadSemanas INT
AS
BEGIN
    SET NOCOUNT ON;

    IF @CantidadSemanas IS NULL OR @CantidadSemanas <= 0
        SET @CantidadSemanas = 1;

    -------------------------------------------------------
    -- 1️⃣ Últimos N LUNES (FechaSemana) desde articost
    -------------------------------------------------------
    ;WITH UltimosLunes AS (
        SELECT TOP (@CantidadSemanas)
            CAST(DATEADD(DAY, 1, ac.fecha) AS date) AS FechaSemana
        FROM DW.dbo.articost ac
        WHERE ac.tipolista = 'S' --lista semanal
        GROUP BY DATEADD(DAY,1, ac.fecha)
        ORDER BY DATEADD(DAY,1, ac.fecha) DESC
    ),

    -------------------------------------------------------
    -- 2️⃣ Agregado comercial semanal (solo se ajusta esta parte)
    -------------------------------------------------------
    ComercialAgg AS (
        SELECT
            dt.CalendarYear,
            dt.WeekNumberOfYear,
            f.COD_ART,

            SUM(
                COALESCE(f.MANO_OBRA_DESPACHO,0)
              + COALESCE(f.COSTO_FLETE,0)
              + COALESCE(f.COSTO_ACARREOS,0)
              + COALESCE(f.COSTO_COMISION_VENTA,0)
              + COALESCE(f.COSTO_COMISION_COBRANZAS,0)
              + COALESCE(f.DESCUENTO_COMISIONES,0)
              + COALESCE(f.COSTO_REPOSITORAS,0)
              + COALESCE(f.ACUERDOS_FIJOS,0)
              + COALESCE(f.FORTALECIMIENTO,0)
              + COALESCE(f.IMPUESTOS,0)
              - CASE 
                    WHEN clasif.ID_Clasificacion_Comprobante = 2 
                    THEN COALESCE(f.IMP_SINIVA,0)
                    ELSE 0
                END
            ) AS CostoComercial,

            -- 👇 KGs solo de clasif 1 (como el SP viejo)
            SUM(
                CASE 
                    WHEN clasif.ID_Clasificacion_Comprobante = 1 THEN COALESCE(f.KGS,0)
                    ELSE 0
                END
            ) AS KgsClasif1

        FROM DW.dbo.FacFacturacion f
        INNER JOIN DW.dbo.DimTime dt
            ON dt.TimeKey = f.clave_fecha_fact
        LEFT JOIN DW.dbo.DimClasificacionComprobante clasif
            ON f.ID_TIPO_COMPROBANTE = clasif.ID_TIPO_COMPROBANTE
        WHERE (f.anulada IS NULL OR f.anulada = 0)
        GROUP BY
            dt.CalendarYear,
            dt.WeekNumberOfYear,
            f.COD_ART
    )

    -------------------------------------------------------
    -- 3️⃣ Resultado final alineado a #CostosBase
    -------------------------------------------------------
    SELECT
        dt.CalendarYear AS Anio,
        dt.WeekNumberOfYear AS Semana,
        DATENAME(MONTH, ul.FechaSemana) AS Mes,
        p.UNIDAD_ME,
        ul.FechaSemana,
        ac.cod_art AS Articulo,
        p.DESCRI_AR,
        cp.GRAN_RUBRO_CDG,
        ta.DESCRIPCION_TIPO_ART,

        -- 10 CostoUnitario
        CASE
            WHEN p.UNIDAD_ME = 1 AND p.COD_PRES = 10 THEN
                ac.COSTO * NULLIF(p.K_BRU_BUL,0)

            WHEN p.UNIDAD_ME = 1 AND p.COD_PRES = 7 THEN
                CASE
                    WHEN p.FORMA_COM = 'C' AND p.CONT_CAJA > 1 THEN
                        ac.COSTO * (p.K_BRU_BUL / NULLIF(p.CONT_CAJA,0))
                    ELSE
                        ac.COSTO * (p.odooKGSCAJA / NULLIF(p.odooCONTCAJA,0))
                END

            ELSE
                ac.COSTO
        END AS CostoUnitario,

        
        -- 11 CostoComercialUnitario
        CASE
            WHEN ca.KgsClasif1 IS NULL OR ca.KgsClasif1 = 0 THEN NULL

            WHEN p.UNIDAD_ME = 1 AND p.COD_PRES = 10 THEN
                (ca.CostoComercial / NULLIF(ca.KgsClasif1,0))
                * NULLIF(p.K_BRU_BUL,0)

            WHEN p.UNIDAD_ME = 1 AND p.COD_PRES = 7 THEN
                (ca.CostoComercial / NULLIF(ca.KgsClasif1,0))
                *
                CASE
                    WHEN p.FORMA_COM = 'C' AND p.CONT_CAJA > 1 THEN
                        (p.K_BRU_BUL / NULLIF(p.CONT_CAJA,0))
                    ELSE
                        (p.odooKGSCAJA / NULLIF(p.odooCONTCAJA,0))
                END

            WHEN p.UNIDAD_ME = 1 THEN
                (ca.CostoComercial / NULLIF(ca.KgsClasif1,0))
                * NULLIF(p.K_BRU_BUL,0)

            ELSE
                (ca.CostoComercial / NULLIF(ca.KgsClasif1,0))
        END AS CostoComercialUnitario

    FROM DW.dbo.articost ac
    INNER JOIN UltimosLunes ul
        ON DATEADD(DAY,1, ac.fecha) = ul.FechaSemana
    LEFT JOIN DW.dbo.DimTime dt
        ON dt.FullDateAlternateKey = ul.FechaSemana
    LEFT JOIN ComercialAgg ca
        ON ca.CalendarYear = dt.CalendarYear
       AND ca.WeekNumberOfYear = dt.WeekNumberOfYear
       AND ca.COD_ART = ac.cod_art
    LEFT JOIN DW.dbo.DimProducto p
        ON p.COD_ART = ac.cod_art
    LEFT JOIN DW.dbo.DimClasificacionProducto cp
        ON cp.COD_ART = ac.cod_art
    LEFT JOIN DW.dbo.DimTipoArticulo ta
        ON p.TIPO_ART = ta.CODIGO_TIPO_ART
    WHERE ac.tipolista = 'S'
    ORDER BY ul.FechaSemana DESC, ac.cod_art;

END;
