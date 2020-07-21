<?php
//call the autoload
require 'vendor/autoload.php';

# Nuestra base de datos
require_once "bd.php";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
//use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

# Obtener base de datos
$bd = obtenerBD();

$documento = new Spreadsheet();
$documento
    ->getProperties()
    ->setCreator("Carlos Sánchez")
    ->setLastModifiedBy('Ingenova')
    ->setTitle('Archivo exportado desde MySQL')
    ->setDescription('Un archivo de Excel exportado desde MySQL por Ingenova');

# Como ya hay una hoja por defecto, la obtenemos, no la creamos
$hojaDeProductos = $documento->getActiveSheet();
$hojaDeProductos->setTitle("Productos");

# Escribir encabezado de los productos
$encabezado = ["Código de barras", "Descripción", "Precio de compra", "Precio de venta", "Existencia"];
# El último argumento es por defecto A1 pero lo pongo para que se explique mejor
$hojaDeProductos->fromArray($encabezado, null, 'A1');

$consulta = "select codigo, descripcion, precioCompra, precioVenta, existencia from productos";
$sentencia = $bd->prepare($consulta, [
    PDO::ATTR_CURSOR => PDO::CURSOR_SCROLL,
]);
$sentencia->execute();
# Comenzamos en la 2 porque la 1 es del encabezado
$numeroDeFila = 2;
while ($producto = $sentencia->fetchObject()) {
    # Obtener los datos de la base de datos
    $codigo = $producto->codigo;
    $descripcion = $producto->descripcion;
    $precioCompra = $producto->precioCompra;
    $precioVenta = $producto->precioVenta;
    $existencia = $producto->existencia;
    # Escribirlos en el documento
    $hojaDeProductos->setCellValueByColumnAndRow(1, $numeroDeFila, $codigo);
    $hojaDeProductos->setCellValueByColumnAndRow(2, $numeroDeFila, $descripcion);
    $hojaDeProductos->setCellValueByColumnAndRow(3, $numeroDeFila, $precioCompra);
    $hojaDeProductos->setCellValueByColumnAndRow(4, $numeroDeFila, $precioVenta);
    $hojaDeProductos->setCellValueByColumnAndRow(5, $numeroDeFila, $existencia);
    $numeroDeFila++;
}

# Ahora los clientes
# Ahora sí creamos una nueva hoja
$hojaDeClientes = $documento->createSheet();
$hojaDeClientes->setTitle("Clientes");

# Escribir encabezado
$encabezado = ["Nombre", "Correo electrónico"];
# El último argumento es por defecto A1 pero lo pongo para que se explique mejor
$hojaDeClientes->fromArray($encabezado, null, 'A1');
# Obtener clientes de BD
$consulta = "select nombre, correo from clientes";
$sentencia = $bd->prepare($consulta, [
    PDO::ATTR_CURSOR => PDO::CURSOR_SCROLL,
]);
$sentencia->execute();

# Comenzamos en la 2 porque la 1 es del encabezado
$numeroDeFila = 2;
while ($cliente = $sentencia->fetchObject()) {
    # Obtener los datos de la base de datos
    $nombre = $cliente->nombre;
    $correo = $cliente->correo;

    # Escribirlos en el documento
    $hojaDeClientes->setCellValueByColumnAndRow(1, $numeroDeFila, $nombre);
    $hojaDeClientes->setCellValueByColumnAndRow(2, $numeroDeFila, $correo);
    $numeroDeFila++;
}
//set the header first, so the result will be treated as an xlsx file.
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

//make it an attachment so we can define filename
header('Content-Disposition: attachment;filename="result.xlsx"');

//create IOFactory object
$writer = IOFactory::createWriter($documento, 'Xlsx');
//save into php output 
$writer->save('php://output');

/*# Crear un "escritor"
$writer = new Xlsx($documento);
# Le pasamos la ruta de guardado
$writer->save('Exportado.xlsx');
echo "<meta http-equiv='refresh' content='0;url=Exportado.xlsx'/>";*/