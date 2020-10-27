CREATE TABLE "DATOSCONEXINET" (
 "id_datosconexinet" IDENTITY DEFAULT '0' NOT NULL,
 "CodAlumbrado" INTEGER UNIQUE,
 "direc_MAC" VARCHAR(50),
 "ubic_fisica" VARCHAR(50),
 "ubic_logica" VARCHAR(50)
)#


-- Esto no hace falta en Todd
-- CREATE UNIQUE NOT MODIFIABLE INDEX "Index_0" ON "DATOSCONEXINET" ( "id_datosconexinet" )#

-- ALTER TABLE "DATOSCONEXINET" 
 -- ADD CONSTRAINT fk_CodAlumbrado
 -- FOREIGN KEY (CodAlumbrado)
 -- REFERENCES ASUMALUM(CodAlumbrado)
 -- ON DELETE CASCADE#
CREATE INDEX "index_2" ON "DATOSCONEXINET" ( "direc_MAC" )#
CREATE INDEX "index_3" ON "DATOSCONEXINET" ( "ubic_fisica" )#
CREATE INDEX "index_4" ON "DATOSCONEXINET" ( "ubic_logica" )#
