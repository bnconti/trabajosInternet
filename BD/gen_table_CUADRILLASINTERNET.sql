CREATE TABLE "CUADRILLASINTERNET" (
 "idcuadrilla" IDENTITY DEFAULT = 1,
 "miembros" VARCHAR(50),
 "email" VARCHAR(100),
 "habilitado" INTEGER
)#
CREATE UNIQUE NOT MODIFIABLE "Index_0" ON "CUADRILLAS" ( "idcuadrilla" )#
CREATE INDEX "Index_1" ON "CUADRILLAS" ( "miembros" )#
