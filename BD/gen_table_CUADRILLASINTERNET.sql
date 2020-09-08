CREATE TABLE "CUADRILLASINTERNET" (
 "idcuadrilla" IDENTITY DEFAULT '1' NOT NULL,
 "miembros" VARCHAR(50) NOT NULL,
 "email" VARCHAR(100),
 "habilitado" BIT NOT NULL
)#

CREATE UNIQUE INDEX "Index_1" ON "CUADRILLASINTERNET" ( "miembros" )#
