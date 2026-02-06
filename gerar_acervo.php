<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

// Libera memória máxima e tempo de execução
ini_set('memory_limit', '-1');
set_time_limit(0);

$arquivosParaLer = glob("*.{xlsx,xls}", GLOB_BRACE);
$acervoMestre = [];

echo "Iniciando consolidacao ultra-leve...\n";

foreach ($arquivosParaLer as $nomeArquivo) {
    if (strpos(strtoupper($nomeArquivo), 'RESULTADO') !== false) continue;
    
    echo "Lendo $nomeArquivo... ";
    
    try {
        $reader = IOFactory::createReaderForFile($nomeArquivo);
        $reader->setReadDataOnly(true);
        $spreadsheet = $reader->load($nomeArquivo);
        $worksheet = $spreadsheet->getActiveSheet();
        
        $colProc = null;
        $colNome = null;
        $count = 0;

        // Lê apenas as 2 primeiras linhas para mapear os cabeçalhos
        foreach ($worksheet->getRowIterator(1, 2) as $row) {
            $cellIterator = $row->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(false);
            foreach ($cellIterator as $cell) {
                $val = mb_strtolower(trim((string)$cell->getValue()));
                $col = $cell->getColumn();
                
                // Mapeia colunas de processo
                if (in_array($val, ['processo', 'número de processo judicial', 'numeroProcesso', 'nº do processo'])) {
                    $colProc = $col;
                }
                // Mapeia colunas de nome
                if (in_array($val, ['procurador', 'responsável', 'responsavel', 'nome'])) {
                    $colNome = $col;
                }
            }
            if ($colProc && $colNome) break;
        }

        if ($colProc && $colNome) {
            // Itera pelas linhas a partir da segunda
            foreach ($worksheet->getRowIterator(2) as $row) {
                $p = trim((string)$worksheet->getCell($colProc . $row->getRowIndex())->getValue());
                $n = trim((string)$worksheet->getCell($colNome . $row->getRowIndex())->getValue());

                // Limpeza: remove pontos, traços e zeros à esquerda
                $pLimpo = ltrim(preg_replace('/[^0-9]/', '', $p), '0');

                if ($pLimpo !== '' && $n !== '' && $n !== '0') {
                    $acervoMestre[$pLimpo] = $n;
                    $count++;
                }
            }
            echo "($count encontrados)\n";
        } else {
            echo "(Colunas nao identificadas - Verifique os nomes na planilha)\n";
        }

        // Limpeza de memória agressiva
        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet);
        gc_collect_cycles();

    } catch (Exception $e) {
        echo "Erro: " . $e->getMessage() . "\n";
    }
}

// Salva o CSV
echo "Salvando acervo.csv com " . count($acervoMestre) . " registros...\n";
$fp = fopen('acervo.csv', 'w');
fputcsv($fp, ['Processo', '','','','','','','','','Procurador'], ';'); 
foreach ($acervoMestre as $proc => $nome) {
    fputcsv($fp, [$proc, '', '', '', '', '', '', '', '', $nome], ';');
}
fclose($fp);
echo "Sucesso!\n";