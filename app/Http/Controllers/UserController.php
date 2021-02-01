<?php

namespace App\Http\Controllers;

class UserController extends Controller
{
    public function cloneBlock()
    {
        // Template processor instance creation
        echo date('H:i:s'), ' Creating new TemplateProcessor instance...';
        $file = storage_path('app/public/word/Sample_23_TemplateBlock.docx');
        $templateProcessor = new \PhpOffice\PhpWord\TemplateProcessor($file);

        // Will clone everything between ${tag} and ${/tag}, the number of times. By default, 1.
        $templateProcessor->cloneBlock('CLONEME', 3);
        //berdasarkan ${CLONINGLIST} and ${/CLONINGLIST} di file template
        $templateProcessor->cloneBlock('CLONINGLIST', 9);
        $templateProcessor->deleteBlock('DELETEME');

        echo date('H:i:s'), ' Saving the result document...';
        $result = storage_path('app/public/word/result_Sample_23_TemplateBlock.docx');
        $templateProcessor->saveAs($result);
    }

    public function export()
    {
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

    }
}
