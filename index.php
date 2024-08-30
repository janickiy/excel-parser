<?php

require_once 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

$file_name = 'КРС.xlsx';
$file_path = sprintf('files/%s', $file_name);

/*$db_config= [
    'host' => 'localhost',
    'dbname' => 'janicky',
    'password' => '456456',
    'username' => 'postgres'
];*/
$db_config = [
    'host' => 'uk1.pgsqlserver.com',
    'dbname' => 'janickiy_report',
    'username' => 'janickiy_report',
    'password' => 'Re3453Y1@'
];

$db_con = sprintf('pgsql:host=%s;dbname=%s', $db_config['host'], $db_config['dbname']);

$db = new PDO($db_con, $db_config['username'], $db_config['password']);

$iterator = static function ($data): Generator {
    yield from new ArrayIterator($data);
};

$processed_data = static function (Worksheet $data) use ($iterator): ?string {
    $key_lists = [];
    $init_data = [];
    foreach ($iterator($data->toArray()) as $item) {
        if (empty($key_lists)) {
            if (is_null($item[0])) break;

            $item = array_map(static fn($_item) => htmlspecialchars((string)$_item), $item);
            $key_lists = $item;
        } else {
            if (empty($key_lists)) break;

            $item = array_map(static fn($_item) => htmlspecialchars((string)$_item), $item);
            $init_data[] = array_combine($key_lists, $item);
        }
    }

    return (empty($init_data))
        ? false
        : json_encode($init_data);
};

$open_file = static function (string $file): ?Spreadsheet {
    $reader = IOFactory::createReader('Xlsx');
    $reader->setReadDataOnly(false);

    return $reader->load($file);
};

$processed = static function (Spreadsheet $spreadsheet) use ($iterator, $processed_data, $db) {
    $sheetNames = [];
    foreach ($iterator($spreadsheet->getSheetNames()) as $item_name) {
        $sheetNames[] = $item_name;
    }

    $sheetCount = $spreadsheet->getSheetCount();
    $i = 0;
    while ($i < $sheetCount) {
        if ($options = $processed_data($spreadsheet->getSheet($i))) {
            $db->query(sprintf(
                "insert into report (sheet, options) values ('%s', '%s')",
                htmlspecialchars($sheetNames[$i]),
                $options
            ));
        }

        $i++;
    }

    return true;
};

try {
    $worksheetData = $open_file($file_path);
    $processed($worksheetData);
    $db_con = null;
} catch (Exception $exception) {
    error_log($exception->getMessage() . PHP_EOL . ' - ' . $exception->getTraceAsString(), 4);
    var_dump($exception);
}