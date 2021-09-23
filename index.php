<!DOCTYPE html>
<html>
<head>
    <title>Read Excel By PHP Excel</title>
</head>
<body>
    <center>
        <h2>Read Excel By PHP Excel</h2>

        <?php
            require_once "ClassesPhpExcel/PHPExcel.php";
            $path = "Test.xlsx";
            $reader = PHPExcel_IOFactory::createReaderForFile($path);
            $excelObj = $reader->load($path);

            $worksheet = $excelObj->getSheet('0');

            $lastRow = $worksheet->getHighestRow();
            $columnCount = $worksheet->getHighestDataColumn();
            $columnCountNumber = PHPExcel_cell::columnIndexFromString($columnCount);

            echo "<table border='1'>";
            echo "<tr><td>NAMA</td><td>KELAS</td></tr>";
            for ($row = 2; $row <= $lastRow; $row++) {
                echo "<tr>";
                for ($col = 0; $col < $columnCountNumber; $col++) {
                    echo "<td>";
                    echo $worksheet->getCell(PHPExcel_Cell::stringFromColumnIndex($col).$row)->getValue();
                    echo "</td>";
                }
                echo "</tr>";
            }
            echo "</table>";

        ?>
    </center>
</body>
</html>
