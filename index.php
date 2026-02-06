<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

// 1. FUNÇÃO PARA CARREGAR O ACERVO DE PROCURADORES
function carregarAcervo() {
    $caminho = 'acervo.csv'; // Nome do arquivo que você deve salvar na pasta
    $lista = [];
    
    if (file_exists($caminho)) {
        $handle = fopen($caminho, "r");
        while (($data = fgetcsv($handle, 1000, ",")) !== FALSE) {
            // No seu arquivo: Coluna 0 é o Processo, Coluna 9 é o Procurador
            $numProcesso = preg_replace('/[^0-9]/', '', $data[0]);
            $nomeProcurador = $data[9] ?? '';
            
            if (!empty($numProcesso)) {
                $lista[$numProcesso] = $nomeProcurador;
            }
        }
        fclose($handle);
    }
    return $lista;
}

// 2. FUNÇÃO PARA FORMATAR NÚMERO CNJ (Tutorial DJE)
function formatarCNJ($numero) {
    $n = preg_replace('/[^0-9]/', '', $numero);
    if (strlen($n) < 20) return $numero;
    return substr($n, 0, 7) . '-' . substr($n, 7, 2) . '.' . substr($n, 9, 4) . '.' . substr($n, 13, 1) . '.' . substr($n, 14, 2) . '.' . substr($n, 16, 4);
}

if (isset($_POST['processar'])) {
    $acervoProcuradores = carregarAcervo();
    $csvUpload = $_FILES['arquivo']['tmp_name'];
    $handle = fopen($csvUpload, "r");
    
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    
    // Cabeçalhos do Excel
    $headers = ['numeroComunicacao', 'numeroProcesso', 'tipoComunicacao', 'dataComunicacao', 'prazo', 'Procurador'];
    $sheet->fromArray($headers, NULL, 'A1');

    $i = 2;
    fgetcsv($handle, 1000, ";"); // Pula cabeçalho do CSV que veio do cartório

    while (($col = fgetcsv($handle, 1000, ";")) !== FALSE) {
        $numProcessoOriginal = $col[1] ?? '';
        $numLimpo = preg_replace('/[^0-9]/', '', $numProcessoOriginal);
        
        // BUSCA NO ACERVO: Se achar o número, coloca o nome. Se não, coloca o aviso em vermelho.
        $procuradorEncontrado = $acervoProcuradores[$numLimpo] ?? 'PROCURADOR NÃO LOCALIZADO';

        $sheet->setCellValue('A' . $i, $col[0] ?? '');
        $sheet->setCellValueExplicit('B' . $i, formatarCNJ($numProcessoOriginal), DataType::TYPE_STRING);
        $sheet->setCellValue('C' . $i, $col[11] ?? '');
        $sheet->setCellValue('D' . $i, $col[12] ?? '');
        $sheet->setCellValue('E' . $i, ($col[19] ?? '') . ' ' . ($col[20] ?? ''));
        $sheet->setCellValue('F' . $i, $procuradorEncontrado);

        // Estilo: Se não localizar, pinta de Vermelho
        if ($procuradorEncontrado === 'PROCURADOR NÃO LOCALIZADO') {
            $sheet->getStyle('F'.$i)->getFont()->getColor()->setARGB('FFFF0000');
        }
        $i++;
    }
    fclose($handle);

    // Ajuste de colunas
    foreach (range('A','F') as $colID) { $sheet->getColumnDimension($colID)->setAutoSize(true); }

    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="DJE_COM_PROCURADORES.xlsx"');
    $writer = new Xlsx($spreadsheet);
    $writer->save('php://output');
    exit;
}
?>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <title>DJE - Localizador de Procuradores</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f4f7f6; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
        .box { background: white; padding: 40px; border-radius: 12px; box-shadow: 0 8px 30px rgba(0,0,0,0.1); text-align: center; max-width: 450px; }
        h2 { color: #2c3e50; margin-bottom: 10px; }
        p { color: #7f8c8d; font-size: 14px; margin-bottom: 25px; }
        .btn { background: #007bff; color: white; border: none; padding: 12px 30px; border-radius: 6px; cursor: pointer; font-weight: 600; width: 100%; }
        .btn:hover { background: #0056b3; }
        input[type="file"] { margin-bottom: 20px; border: 1px solid #ddd; padding: 10px; width: 100%; border-radius: 5px; }
    </style>
</head>
<body>
    <div class="box">
        <h2>Processador DJE</h2>
        <p>Certifique-se de que o arquivo <b>acervo.csv</b> está na pasta do XAMPP para localizar os procuradores.</p>
        <form method="post" enctype="multipart/form-data">
            <input type="file" name="arquivo" accept=".csv" required>
            <button type="submit" name="processar" class="btn">Processar e Gerar Excel</button>
        </form>
    </div>
</body>
</html>