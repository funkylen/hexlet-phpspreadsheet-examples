<?php

namespace App;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

const PATH = __DIR__ . '/../excel.xlsx';
const PATH1 = __DIR__ . '/../downloaded.xlsx';
const PATH2 = __DIR__ . '/../csv.csv';

const ARR = [
    ['col1', 'col2', '_col3'],
    ['row1', 'row1', 'row1'],
    ['row2', 'row2', 'row2'],
    ['row3', 'row3', 'row3'],
    ['row4', 'row4', 'row4'],
];

function run()
{
    echo 'hello!'
}

function createSheet()
{
    $spreadsheet = IOFactory::load(PATH);

    $worksheet = $spreadsheet->createSheet(1);
    $worksheet->setTitle('PHP');

    $worksheet->setCellValue('B2', 'PHP 8.1');

    $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
    $writer->save(PATH);
}

function readCustomSheet()
{
    $spreadsheet = IOFactory::load(PATH);
    $worksheet = $spreadsheet->getSheet(1);
    var_dump($spreadsheet->getSheetNames(), $worksheet->getTitle(), $worksheet->getCell('B2')->getValue());
}

function useWriter()
{
    $spreadsheet = new Spreadsheet();
    $activeWorksheet = $spreadsheet->getActiveSheet();

    $activeWorksheet->fromArray(ARR);

    $writer = IOFactory::createWriter($spreadsheet, 'Csv');
    $writer->save(PATH2);
}

function useReader()
{
    $reader = IOFactory::createReaderForFile(PATH1);
    $reader->setReadEmptyCells(false);
    $spreadsheet = $reader->load(PATH1, $reader::READ_DATA_ONLY | $reader::IGNORE_EMPTY_CELLS);
    $worksheet = $spreadsheet->getActiveSheet();
    var_dump($worksheet->toArray());
}

function readSheet()
{
    $spreadsheet = IOFactory::load(PATH1);
    $worksheet = $spreadsheet->getActiveSheet();
    var_dump($worksheet->toArray());
    var_dump($worksheet->rangeToArray('A1:H101'));
}

function writeMatrixSimple()
{
    $spreadsheet = new Spreadsheet();
    $activeWorksheet = $spreadsheet->getActiveSheet();

    $activeWorksheet->fromArray(ARR);

    $writer = new Xlsx($spreadsheet);
    $writer->save(PATH);
}

function writeMatrix()
{
    $spreadsheet = new Spreadsheet();
    $activeWorksheet = $spreadsheet->getActiveSheet();

    $alphabet = range('A', 'Z');

    foreach (ARR as $idx1 => $row) {
        foreach ($row as $idx2 => $value) {
            $letter = $alphabet[$idx2];
            $activeWorksheet->setCellValue($letter . $idx1 + 1, $value);
        }
    }

    $writer = new Xlsx($spreadsheet);
    $writer->save(PATH);
}


function create()
{
    $spreadsheet = new Spreadsheet();
    $activeWorksheet = $spreadsheet->getActiveSheet();

    $activeWorksheet->setCellValue('A1', 'Hello World !');
    $activeWorksheet->setCellValue('C3', 'World !');
    $activeWorksheet->setCellValue('H10', 'Hello !');

    $writer = new Xlsx($spreadsheet);
    $writer->save(PATH);
}
