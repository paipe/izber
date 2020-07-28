<?php

require 'vendor/autoload.php';

$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('izber.xls');
$activeSheet = $spreadsheet->getSheetByName('Приложение 2А');

$columnName = 'H';
$startRow = 3;
$finishRow = 114;
$client = new \GuzzleHttp\Client();
for ($i = $startRow; $i <= $finishRow; $i++) {
    $cellValue = $activeSheet->getCell($columnName . $i)->getValue();
    var_dump($cellValue);

    $response = $client->post(
        'https://suggestions.dadata.ru/suggestions/api/4_1/rs/suggest/address',
        [
            'body'    => sprintf('{ "query": "%s", "count": 1 }', $cellValue),
            'headers' => [
                'Content-Type'  => 'application/json',
                'Accept'        => 'application/json',
                'Authorization' => 'Token 5514707bc777c711f8d68107ecbc97734c543ccd',
            ]
        ]
    );

    $resultData = json_decode((string)$response->getBody(), true)['suggestions'][0]['data'];

    $activeSheet->setCellValue('H' . $i, $resultData['region_with_type']);
    $activeSheet->setCellValue('I' . $i, $resultData['area_with_type']);
    $activeSheet->setCellValue('J' . $i, $resultData['city_with_type']);
    $activeSheet->setCellValue('K' . $i, $resultData['settlement_with_type']);
    $activeSheet->setCellValue('L' . $i, $resultData['street'] . ' ' . $resultData['street_type']);
    $activeSheet->setCellValue('M' . $i, $resultData['house']);
    $activeSheet->setCellValue('N' . $i, $resultData['block']);
    $activeSheet->setCellValue('O' . $i, $resultData['flat']);
}

$writer = new \PhpOffice\PhpSpreadsheet\Writer\Xls($spreadsheet);
$writer->save('test.xls');