<?php
require_once 'PHPExcel/PHPExcel.php';

// Ruta y nombre del archivo de Excel
$excelFile = 'datos_mascotas.xlsx';

// Crear un nuevo objeto PHPExcel
$objPHPExcel = new PHPExcel();

// Seleccionar la hoja activa
$objPHPExcel->setActiveSheetIndex(0);
$sheet = $objPHPExcel->getActiveSheet();

// Obtener los datos del formulario
$nombre = $_POST['nombre'];
$raza = $_POST['raza'];
$genero = $_POST['genero'];
$fecha_perdida = $_POST['fecha_perdida'];
$especie = $_POST['especie'];
$color = $_POST['color'];
$descripcion = $_POST['descripcion'];
$nombre_dueno = $_POST['nombre_dueno'];
$telefono = $_POST['telefono'];
$correo = $_POST['correo'];
$mensaje = $_POST['mensaje'];

// Configurar las celdas con los datos
$sheet->setCellValue('A1', 'Nombre');
$sheet->setCellValue('B1', 'Raza');
$sheet->setCellValue('C1', 'Género');
$sheet->setCellValue('D1', 'Fecha de pérdida');
$sheet->setCellValue('E1', 'Especie');
$sheet->setCellValue('F1', 'Color');
$sheet->setCellValue('G1', 'Descripción');
$sheet->setCellValue('H1', 'Nombre del dueño');
$sheet->setCellValue('I1', 'Teléfono');
$sheet->setCellValue('J1', 'Correo');
$sheet->setCellValue('K1', 'Mensaje');

$sheet->setCellValue('A2', $nombre);
$sheet->setCellValue('B2', $raza);
$sheet->setCellValue('C2', $genero);
$sheet->setCellValue('D2', $fecha_perdida);
$sheet->setCellValue('E2', $especie);
$sheet->setCellValue('F2', $color);
$sheet->setCellValue('G2', $descripcion);
$sheet->setCellValue('H2', $nombre_dueno);
$sheet->setCellValue('I2', $telefono);
$sheet->setCellValue('J2', $correo);
$sheet->setCellValue('K2', $mensaje);

// Guardar el archivo de Excel
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save($excelFile);

// Redireccionar o mostrar un mensaje de éxito
echo "Datos guardados en el archivo de Excel.";
?>
