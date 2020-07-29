<?php

require 'vendor/autoload.php';

$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('out.xls');
$resultSheet = $spreadsheet->getSheetByName('Приложение 2А');

try {
    $columnName = 'H';
    $startRow = 3;
    $finishRow = 82;
    $client = new \GuzzleHttp\Client();
    for ($i = $startRow; $i <= $finishRow; $i++) {
        $finishValue = ($i === $finishRow) ? '1' : '0';
        if ($finishValue === '1') {
            $resultSheet->setCellValue('A' . $i, $i);
            $resultSheet->setCellValue('B' . $i, $finishValue);
        }
        $dateMoneyInput = $resultSheet->getCell('B' . $i)->getValue();
        $fio = explode(' ', $resultSheet->getCell('C' . $i)->getValue());
        $lastName = $fio[0];
        $fatherName = $fio[2];
        $firstName = $fio[1];
        $birthday = $resultSheet->getCell('D' . $i)->getValue();
        $fullAddress = str_replace("\n", '', $resultSheet->getCell('E' . $i)->getValue());
        $documentType = '21';
        $documentData = $resultSheet->getCell('F' . $i)->getValue();
        // Паспорт гражданина РФ 4013 № 790982
        preg_match('/^[^0-9]+([^№]+)№(.+)$/i', $documentData, $matches);
        $serial = str_replace(' ', '', $matches[1]);
        $number = str_replace(' ', '', $matches[2]);
        $country = 'Россия';
        $amount = $resultSheet->getCell('I' . $i)->getValue();

        $response = $client->post(
            'https://suggestions.dadata.ru/suggestions/api/4_1/rs/suggest/address',
            [
                'body'    => sprintf('{ "query": "%s", "count": 1 }', $fullAddress),
                'headers' => [
                    'Content-Type'  => 'application/json',
                    'Accept'        => 'application/json',
                    'Authorization' => 'Token 5514707bc777c711f8d68107ecbc97734c543ccd',
                ]
            ]
        );

        $resultData = json_decode((string)$response->getBody(), true)['suggestions'][0]['data'];

        $resultSheet->setCellValue('H' . $i, $resultData['region_with_type']);
        $resultSheet->setCellValue('I' . $i, $resultData['area_with_type']);
        $resultSheet->setCellValue('J' . $i, $resultData['city_with_type']);
        $resultSheet->setCellValue('K' . $i, $resultData['settlement_with_type']);
        $resultSheet->setCellValue('L' . $i, $resultData['street'] . ' ' . $resultData['street_type']);
        $resultSheet->setCellValue('M' . $i, $resultData['house']);
        $resultSheet->setCellValue('N' . $i, $resultData['block']);
        $resultSheet->setCellValue('O' . $i, $resultData['flat']);

        $resultSheet->setCellValue('B' . $i, $finishValue);
        $resultSheet->setCellValue('C' . $i, $dateMoneyInput);
        $resultSheet->setCellValue('D' . $i, $lastName);
        $resultSheet->setCellValue('E' . $i, $firstName);
        $resultSheet->setCellValue('F' . $i, $fatherName);
        $resultSheet->setCellValue('G' . $i, $birthday);
        $resultSheet->setCellValue('P' . $i, $documentType);
        $resultSheet->setCellValue('Q' . $i, $serial);
        $resultSheet->setCellValue('R' . $i, $number);
        $resultSheet->setCellValue('S' . $i, $country);
        $resultSheet->setCellValue('U' . $i, $amount);

        $resultSheet->getStyle('B' . $i)
            ->getNumberFormat()
            ->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER);
        $resultSheet->getStyle('C' . $i)
            ->getNumberFormat()
            ->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
        $resultSheet->getStyle('G' . $i)
            ->getNumberFormat()
            ->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);

    }
} finally {
    $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xls($spreadsheet);
    $writer->save('result.xls');
}

