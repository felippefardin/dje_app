<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;

// 1. CARREGAR ACERVO COM MAPEAMENTO CORRETO
function carregarAcervo() {
    $caminho = 'acervo.csv'; 
    $lista = [];
    
    if (file_exists($caminho)) {
        $handle = fopen($caminho, "r");
        // Pula o cabeçalho do acervo
        fgetcsv($handle, 2000, ","); 

        while (($data = fgetcsv($handle, 2000, ",")) !== FALSE) {
            // Coluna 0: Processo | Coluna 9: Procurador (conforme seu arquivo acervo.csv)
            $numProcesso = preg_replace('/[^0-9]/', '', $data[0]);
            $nomeProcurador = trim($data[9] ?? '');
            
            if (!empty($numProcesso) && $nomeProcurador !== 'NaN' && !empty($nomeProcurador)) {
                $lista[$numProcesso] = $nomeProcurador;
            }
        }
        fclose($handle);
    }
    return $lista;
}

// 2. FORMATAÇÃO CNJ
function formatarCNJ($numero) {
    $n = preg_replace('/[^0-9]/', '', $numero);
    if (strlen($n) < 15) return $numero; // Retorna original se for muito curto
    // Tenta formatar se tiver os 20 dígitos padrão
    if (strlen($n) == 20) {
        return substr($n, 0, 7) . '-' . substr($n, 7, 2) . '.' . substr($n, 9, 4) . '.' . substr($n, 13, 1) . '.' . substr($n, 14, 2) . '.' . substr($n, 16, 4);
    }
    return $numero;
}

if (isset($_POST['processar']) && isset($_FILES['arquivo'])) {
    $acervoProcuradores = carregarAcervo();
    $csvUpload = $_FILES['arquivo']['tmp_name'];
    $handle = fopen($csvUpload, "r");
    
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    
    // CABEÇALHOS NO MODELO "RESULTADO"
    $headers = [
        'Processo', 'Processo', 'Tipo de Comunicacao', 'Data da Comunicacao', 
        'Data Final p/ Ciencia', 'Prazo', 'Tipo de Prazo', 'Data da Ciência', 
        'Ciência Automática?', 'Procurador'
    ];
    $sheet->fromArray($headers, NULL, 'A1');

    // Estilo do Cabeçalho
    $sheet->getStyle('A1:J1')->getFont()->setBold(true);
    $sheet->getStyle('A1:J1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setRGB('E0E0E0');

    $i = 2;
    fgetcsv($handle, 1000, ";"); // Pula cabeçalho do arquivo enviado

    while (($col = fgetcsv($handle, 1000, ";")) !== FALSE) {
        if (empty($col[1])) continue;

        $numProcessoOriginal = $col[1];
        $numLimpo = preg_replace('/[^0-9]/', '', $numProcessoOriginal);
        
        // Busca o procurador no acervo carregado
        $procuradorEncontrado = $acervoProcuradores[$numLimpo] ?? 'NÃO LOCALIZADO NO ACERVO';

        // Mapeamento baseado nas colunas do CSV do cartório para o modelo RESULTADO
        $sheet->setCellValueExplicit('A' . $i, formatarCNJ($numProcessoOriginal), DataType::TYPE_STRING); // Coluna A: Processo formatado
        $sheet->setCellValue('B' . $i, ''); // Coluna B: Vazia conforme modelo
        $sheet->setCellValue('C' . $i, $col[11] ?? ''); // Tipo de Comunicação
        $sheet->setCellValue('D' . $i, $col[12] ?? ''); // Data Comunicação
        $sheet->setCellValue('E' . $i, $col[13] ?? ''); // Data Final Ciência
        $sheet->setCellValue('F' . $i, $col[19] ?? ''); // Prazo
        $sheet->setCellValue('G' . $i, $col[20] ?? ''); // Tipo de Prazo (Dias)
        $sheet->setCellValue('H' . $i, ''); // Data da Ciência
        $sheet->setCellValue('I' . $i, 'S'); // Ciência Automática
        $sheet->setCellValue('J' . $i, $procuradorEncontrado); // Procurador Identificado

        // Destaque se não localizar
        if ($procuradorEncontrado === 'NÃO LOCALIZADO NO ACERVO') {
            $sheet->getStyle('J'.$i)->getFont()->getColor()->setARGB('FFFF0000');
        }
        $i++;
    }
    fclose($handle);

    // Ajuste automático das colunas
    foreach (range('A','J') as $colID) { $sheet->getColumnDimension($colID)->setAutoSize(true); }

    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="RESULTADO_PROCESSADO.xlsx"');
    
    $writer = new Xlsx($spreadsheet);
    $writer->save('php://output');
    exit;
}
?>