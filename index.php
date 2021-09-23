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

            echo $worksheet->getCell('A2')->getValue();

            $lastRow = $worksheet->getHighestRow();
            $columnCount = $worksheet->getHighestDataColumn();

            echo $lastRow;
            echo $columnCount;
        ?>

    </center>
</body>
</html>
