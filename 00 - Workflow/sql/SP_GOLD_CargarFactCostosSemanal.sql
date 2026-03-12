USE [DW]
GO
/****** Object:  StoredProcedure [dbo].[SP_GOLD_CargarFactCostosSemanal]    Script Date: 11/03/2026 12:03:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [dbo].[SP_GOLD_CargarFactCostosSemanal]
AS
BEGIN

SET NOCOUNT ON;

-------------------------------------------------------
-- VARIABLES
-------------------------------------------------------

DECLARE @FechaHoy DATE = CAST(GETDATE() AS DATE);

-------------------------------------------------------
-- LIMPIAR TABLA DESTINO
-------------------------------------------------------

TRUNCATE TABLE GOLD.FactCostosSemanal;

-------------------------------------------------------
-- 1️⃣ COSTOS UNITARIOS
-------------------------------------------------------

CREATE TABLE #CostosUnitarios
(
    CalendarYear INT,
    WeekNumberOfYear INT,
    FechaSemana DATE,
    Articulo INT,
    UNIDAD_ME INT,
    DESCRI_AR VARCHAR(255),
    GRAN_RUBRO_CDG VARCHAR(50),
    DESCRIPCION_TIPO_ART VARCHAR(100),
    CostoUnitario DECIMAL(18,4)
);

INSERT INTO #CostosUnitarios
EXEC dbo.SP_ListarCostosUnitarios @CantidadSemanas = 260;

CREATE INDEX IX_CostosUnitarios
ON #CostosUnitarios (Articulo, FechaSemana);

-------------------------------------------------------
-- 2️⃣ COSTOS COMERCIALES
-------------------------------------------------------

CREATE TABLE #CostosComerciales
(
    CalendarYear INT,
    WeekNumberOfYear INT,
    COD_ART INT,

    ManoObraDespacho FLOAT,
    CostoFlete FLOAT,
    CostoAcarreos FLOAT,
    ComisionVentas FLOAT,
    ComisionCobranzas FLOAT,
    DescuentoComisiones FLOAT,
    CostoRepositoras FLOAT,
    AcuerdosFijos FLOAT,
    Fortalecimiento FLOAT,
    Impuestos FLOAT,
    NcFinancieras FLOAT,

    CostoComercial DECIMAL(18,4),
    KgsClasif1 FLOAT,
    CostoComercialUnitario DECIMAL(18,4)
);

INSERT INTO #CostosComerciales
EXEC dbo.SP_ListarCostosComerciales @CantidadSemanas = 260;

CREATE INDEX IX_CostosComerciales
ON #CostosComerciales (COD_ART, CalendarYear, WeekNumberOfYear);

-------------------------------------------------------
-- 3️⃣ BASE UNIFICADA
-------------------------------------------------------

CREATE TABLE #CostosBase
(
    Anio INT,
    Semana INT,
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
SELECT
    u.CalendarYear,
    u.WeekNumberOfYear,
    u.UNIDAD_ME,
    u.FechaSemana,
    u.Articulo,
    u.DESCRI_AR,
    u.GRAN_RUBRO_CDG,
    u.DESCRIPCION_TIPO_ART,
    u.CostoUnitario,
    c.CostoComercialUnitario
FROM #CostosUnitarios u
LEFT JOIN #CostosComerciales c
       ON c.CalendarYear = u.CalendarYear
      AND c.WeekNumberOfYear = u.WeekNumberOfYear
      AND c.COD_ART = u.Articulo;

CREATE INDEX IX_CostosBase
ON #CostosBase (Articulo, FechaSemana);

-------------------------------------------------------
-- 4️⃣ SEMANA ACTUAL
-------------------------------------------------------

DECLARE @FechaFinSemanaActual DATE;

SELECT
    @FechaFinSemanaActual = MAX(FullDateAlternateKey)
FROM DW.dbo.DimTime
WHERE CalendarYear = (
        SELECT CalendarYear
        FROM DW.dbo.DimTime
        WHERE FullDateAlternateKey = @FechaHoy
)
AND WeekNumberOfYear = (
        SELECT WeekNumberOfYear
        FROM DW.dbo.DimTime
        WHERE FullDateAlternateKey = @FechaHoy
);

-------------------------------------------------------
-- 5️⃣ RANGO INICIAL POR ARTICULO
-------------------------------------------------------

;WITH RangoInicialArticulo AS
(
    SELECT
        Articulo,
        MIN(FechaSemana) AS FechaMinEvento
    FROM #CostosBase
    GROUP BY Articulo
),

-------------------------------------------------------
-- 6️⃣ CALENDARIO SEMANAL
-------------------------------------------------------

Semanas AS
(
    SELECT
        r.Articulo,
        dt.CalendarYear,
        dt.WeekNumberOfYear,
        MIN(dt.FullDateAlternateKey) AS FechaSemana,
        MAX(dt.FullDateAlternateKey) AS FechaFinSemana
    FROM RangoInicialArticulo r
    INNER JOIN DW.dbo.DimTime dt
        ON dt.FullDateAlternateKey BETWEEN r.FechaMinEvento
                                       AND @FechaFinSemanaActual
    GROUP BY
        r.Articulo,
        dt.CalendarYear,
        dt.WeekNumberOfYear
)

-------------------------------------------------------
-- 7️⃣ FILL FORWARD
-------------------------------------------------------

INSERT INTO GOLD.FactCostosSemanal
(
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

    u.DESCRI_AR,
    u.GRAN_RUBRO_CDG,
    u.UNIDAD_ME,

    u.CostoUnitario,
    cc.CostoComercialUnitario

FROM Semanas s

-------------------------------------------------------
-- UNITARIO (fill forward)
-------------------------------------------------------

OUTER APPLY
(
    SELECT TOP 1
        cb.CostoUnitario,
        cb.DESCRI_AR,
        cb.GRAN_RUBRO_CDG,
        cb.UNIDAD_ME
    FROM #CostosBase cb
    WHERE cb.Articulo = s.Articulo
      AND cb.FechaSemana <= s.FechaFinSemana
      AND cb.CostoUnitario IS NOT NULL
      AND cb.CostoUnitario <> 0
    ORDER BY cb.FechaSemana DESC
) u

-------------------------------------------------------
-- COMERCIAL (fill forward independiente)
-------------------------------------------------------

OUTER APPLY
(
    SELECT TOP 1
        cb.CostoComercialUnitario
    FROM #CostosBase cb
    WHERE cb.Articulo = s.Articulo
      AND cb.FechaSemana <= s.FechaFinSemana
      AND cb.CostoComercialUnitario IS NOT NULL
      AND cb.CostoComercialUnitario <> 0
    ORDER BY cb.FechaSemana DESC
) cc

-------------------------------------------------------
-- REGLA DE NEGOCIO
-------------------------------------------------------

WHERE u.CostoUnitario IS NOT NULL;

-------------------------------------------------------
-- LIMPIEZA
-------------------------------------------------------

DROP TABLE #CostosUnitarios;
DROP TABLE #CostosComerciales;
DROP TABLE #CostosBase;

END;

