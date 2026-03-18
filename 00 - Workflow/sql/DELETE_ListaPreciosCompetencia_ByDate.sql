USE [DW]
GO

-- Verificar registros a eliminar
SELECT * FROM dbo.Lista_Precios_Competencia
WHERE CAST(fecha_carga AS DATE) = '2026-03-16'

-- Eliminar registros de Lista_Precios_Competencia por fecha de carga
DELETE FROM dbo.Lista_Precios_Competencia
WHERE CAST(fecha_carga AS DATE) = '2026-03-16'
