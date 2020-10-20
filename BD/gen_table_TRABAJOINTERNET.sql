CREATE TABLE "TRABAJOINTERNET" (
 "id_trabajo" IDENTITY DEFAULT '0' NOT NULL,
 "nroorden" INTEGER NOT NULL,
 "estado" INTEGER NOT NULL,
 "fecha_pedido" DATE,
 "fecha_inst" DATE NOT NULL,
 "hora_inst" TIME,
 "tipo_conexion" VARCHAR(30),
 "idcuadrilla" INTEGER,
 "ancho_banda" VARCHAR(50),
 "prioridad" INTEGER,
 "obs" VARCHAR(50),
 "reserva" VARCHAR(50)
)#


-- Esto hace falta en algunas versiones de Pervasive, sino no te pone el ID como autoincremental. Y en otras no hace falta y si lo pones te genera error.
--CREATE UNIQUE NOT MODIFIABLE INDEX "index_0" ON "TRABAJOINTERNET" ( "id_trabajo" )#
CREATE INDEX "index_1" ON "TRABAJOINTERNET" ( "nroorden" )#
CREATE INDEX "index_2" ON "TRABAJOINTERNET" ( "estado" )#
CREATE INDEX "index_3" ON "TRABAJOINTERNET" ( "fecha_pedido" )#
CREATE INDEX "index_4" ON "TRABAJOINTERNET" ( "fecha_inst" )#
CREATE INDEX "index_5" ON "TRABAJOINTERNET" ( "idcuadrilla" )#
CREATE INDEX "index_6" ON "TRABAJOINTERNET" ( "prioridad" )#

-- Si es la tabla vieja que no tiene estos campos
ALTER TABLE "TRABAJOINTERNET" ADD ancho_banda VARCHAR(50)#
ALTER TABLE "TRABAJOINTERNET" ADD prioridad INTEGER#