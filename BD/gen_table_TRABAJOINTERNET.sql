CREATE TABLE "TRABAJOINTERNET" (
 "id_trabajo" IDENTITY DEFAULT '0' NOT NULL,
 "nroorden" INTEGER NOT NULL,
 "estado" INTEGER NOT NULL,
 "fecha_pedido" DATE,
 "fecha_inst" DATE,
 "hora_inst" TIME,
 "tipo_conexion" VARCHAR(30),
 "idcuadrilla" INTEGER,
 "obs" VARCHAR(50),
 "reserva" VARCHAR(50)
)#

CREATE UNIQUE NOT MODIFIABLE INDEX "index_0" ON "TRABAJOINTERNET" ( "id_trabajo" )#
CREATE INDEX "index_1" ON "TRABAJOINTERNET" ( "nroorden" )#
CREATE INDEX "index_2" ON "TRABAJOINTERNET" ( "estado" )#
CREATE INDEX "index_3" ON "TRABAJOINTERNET" ( "fecha_pedido" )#
CREATE INDEX "index_4" ON "TRABAJOINTERNET" ( "fecha_inst" )#
CREATE INDEX "index_5" ON "TRABAJOINTERNET" ( "idcuadrilla" )#
