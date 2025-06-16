--PRIMERO TABLA: riesgoPrefijos

--PASO 1. Agregar Campo Días
--Nota: El Default actualiza inmediatamente los registros existentes en 1.

ALTER TABLE riesgoPrefijos ADD Dias Int NOT NULL DEFAULT(1)

--PASO 2. Recrear Indice

CREATE UNIQUE CLUSTERED INDEX [riesgoPrefijosIDX_01] ON [dbo].[riesgoPrefijos]([limInf], [limSup], [mcc], [pais], [flag2], [Dias]) 
WITH  FILLFACTOR = 100,DROP_EXISTING ON [PRIMARY]

--PASO 3. Actualizar Flag 3 a S que indica que esta Habilitado el parametro.
--De esta forma todos los Parametros actuales siguen habilitados.

UPDATE riesgoPrefijos set flag3='S'

--PASO 4. RECREAR VISTA dbo.vista_riesgoPrefijos.VIW

____________________________________________________


--SEGUNDO TABLA: log_riesgoPrefijos

--PASO 1. Agregar Campo Días
--Nota: El Default actualiza inmediatamente los registros existentes en 1.

ALTER TABLE log_riesgoPrefijos ADD fecha datetime NULL DEFAULT(getdate())
ALTER TABLE log_riesgoPrefijos ADD Dias Int NOT NULL DEFAULT(1)

--PASO 2. Crear Indices

 CREATE  INDEX [log_riesgoPrefijosIDX_01] ON [dbo].[log_riesgoPrefijos]([limInf], [limSup], [mcc], [pais], [flag2], [Dias], [tiempo]) WITH  FILLFACTOR = 100 ON [PRIMARY]
GO

 CREATE  INDEX [log_riesgoPrefijosIDX_02] ON [dbo].[log_riesgoPrefijos]([limInf]) WITH  FILLFACTOR = 80 ON [PRIMARY]
GO

--PASO 3. Actualizar Flag 3 a S que indica que esta Habilitado el parametro.
--De esta forma todos los Parametros actuales siguen habilitados.

UPDATE log_riesgoPrefijos set flag3='S'

--PASO 4. RECREAR VISTA dbo.Vista_Log_riesgoPrefijos.VIW