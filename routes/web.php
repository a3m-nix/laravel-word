<?php

use App\Models\User;
use Illuminate\Support\Facades\Route;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
 */
Route::get('users/word', function () {
    $phpWord = new \PhpOffice\PhpWord\PhpWord();
    $section = $phpWord->addSection();
    $styleHeader =
        [
        'bgColor' => 'd0cece',
        'borderColor' => '000000',
        'borderSize' => 6,
    ];
    $styleCell =
        [
        'borderColor' => '000000',
        'borderSize' => 6,
    ];

    $header = array('size' => 16, 'bold' => true);
    $section->addText(htmlspecialchars('Basic table'), $header);

    $table = $section->addTable();
    //header table
    $table->addRow();
    //parameter null ukuran width kolom
    $table->addCell(1000, $styleHeader)->addText("No");
    $table->addCell(null, $styleHeader)->addText("NAME");
    $table->addCell(null, $styleHeader)->addText("EMAIL");
    $users = User::latest()->get();
    $i = 1;
    foreach ($users as $item) {
        //isi tabel
        $table->addRow();
        $table->addCell(null, $styleCell)->addText($i);
        $table->addCell(null, $styleCell)->addText($item->name);
        $table->addCell(null, $styleCell)->addText($item->email);
        $i++;
    }

    $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
    $objWriter->save(storage_path('helloWorld.docx'));
    return response()->download(storage_path('helloWorld.docx'));
});

Route::get('/', function () {
    return view('welcome');
});
