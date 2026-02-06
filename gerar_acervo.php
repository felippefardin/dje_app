<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

$arquivosParaLer = [
    'PJE-Busca ACERVO.xlsx', 
    'data.xlsx',
    'DJEBUSCA.xlsx'
];

$acervoMestre = [];

echo "Iniciando consolidação do acervo...\n";

foreach ($arquivosParaLer as $nomeArquivo) {
    if (!file_exists($nomeArquivo)) {
        echo "Aviso: Arquivo $nomeArquivo não encontrado. Pulando...\n";
        continue;
    }

    echo "Lendo $nomeArquivo...\n";
    $spreadsheet = IOFactory::load($nomeArquivo);
    $rows = $spreadsheet->getActiveSheet()->toArray(null, true, true, true);
    $header = array_shift($rows);

    // Identifica as colunas de Processo e Procurador/Responsável
    $colProc = 'A';
    $colNome = 'J'; // Padrão do seu PJE-Busca

    foreach ($header as $col => $val) {
        $val = mb_strtolower(trim($val));
        if ($val == 'processo' || $val == 'número de processo judicial') $colProc = $col;
        if ($val == 'procurador' || $val == 'responsável') $colNome = $col;
    }

    foreach ($rows as $row) {
        $proc = preg_replace('/[^0-9]/', '', (string)$row[$colProc]);
        $nome = trim($row[$colNome] ?? '');

        if ($proc && $nome && $nome !== '0') {
            $acervoMestre[$proc] = $nome;
        }
    }
}

// Salva o arquivo CSV final
$fp = fopen('acervo.csv', 'w');
fputcsv($fp, ['Processo', '', '', '', '', '', '', '', '', 'Procurador'], ';'); // Cabeçalho Fake para manter compatibilidade

foreach ($acervoMestre as $proc => $nome) {
    // Escreve o processo na col 0 e o nome na col 9 (índice J)
    $linha = [$proc, '', '', '', '', '', '', '', '', $nome];
    fputcsv($fp, $linha, ';');
}

fclose($fp);
echo "Pronto! Arquivo 'acervo.csv' gerado com sucesso com " . count($acervoMestre) . " processos mapeados.\n";