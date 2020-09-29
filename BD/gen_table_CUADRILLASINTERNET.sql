CREATE TABLE "CUADRILLASINTERNET" (
 "idcuadrilla" IDENTITY DEFAULT '0' NOT NULL,
 "miembros" VARCHAR(50) NOT NULL,
 "email" VARCHAR(100),
 "habilitado" BIT NOT NULL
)#

--CREATE UNIQUE NOT MODIFIABLE INDEX "Index_0" ON "CUADRILLASINTERNET" ( "idcuadrilla" )#
CREATE UNIQUE INDEX "Index_1" ON "CUADRILLASINTERNET" ( "miembros" )#
