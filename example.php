<?php
error_reporting(E_ALL);
if(ini_set('display_errors', 1)===false)
  echo "ERROR INI SET";

require 'vendor/autoload.php';
require_once 'docwork.class.php';

$rep = new DocWork(['orientation' => 'landscape']);

$html
= "<div>
  <div style='color:#FF0000;font-size:20px;font-weight:bold;'>Информация</div>
  <p><b>Name:</b> Ivanoff Kolya</p>
  <p><b>Date:</b> ".date('Y-m-d')."</p>
</div>";

$rep
  ->html($html)
  ->title('Инкапсуляция, полиморфизм, объектное мышление…?', 1)
  ->text('Brave New World', ['size' => 22, 'color' => '00FF00', 'underline' => 'single']);

$arr = [
  [1, '',     'Системный блок', '',                             21840, 5, 109200],
  [2, 295158, 'Монитор',        'LG 22" 22MK400H-B22',          6460,  3, 19380 ],
  [3, 7006,   'SSD диск',       '240Gb SSD Kingston A400',      2610,  5, 13050 ],
  [4, 17706,  'Свитч',          'D-Link DGS-1008D',             1710,  2, 3420  ],
  [5, 54759,  'Гарнитура',      'Logitech Stereo Headset H110', 1200,  2, 2400  ]
];

$head = ['№ пп', 'ID', 'Наименование', 'Модель', 'Цена, руб.', 'Кол-во', 'Сумма, руб.'];
$width= [ 1,     2,    4,              6,         2,           2,        2];

$options = [
  'head' => $head,
  'width' => $width,
  'caption' => 'Закупка техники',
  'fontSize' => 12,
  'columns' => [
    ['bold' => true, 'alignment' => 'center'],
    ['alignment' => 'right'],
    [],
    ['color' => 'AAAAAA', 'italic' => true],
    ['alignment' => 'right'],
    ['alignment' => 'right'],
    ['alignment' => 'right'],
  ]
];

$rep
  ->table($arr, $options)
  ->textBreak(3)
  ->header('Simple Report!!!', ['bold' => true])
  ->pageNumber(false);

$rep
  ->title('Test', 10)
  ->title('Test', 11)
  ->title('Test', 11)
  ->title('Test', 12)
  ->title('Test', 12)
  ->title('Test', 10)
  ->title('Test', 11)
  ->title('Test', 12)
  ->title('Test', 12);

$rep->save();
