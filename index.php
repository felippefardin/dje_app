<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

/* ===============================
   CARREGAR ACERVO (BASE DE DADOS)
================================ */
function carregarAcervo() {
    $lista = [];
    $caminho = __DIR__ . '/acervo.csv';
    if (!file_exists($caminho)) return $lista;

    // Aumentamos o limite de memória para acervos grandes
    ini_set('memory_limit', '512M');

    $h = fopen($caminho, 'r');
    fgetcsv($h, 2000, ';'); 

    while (($d = fgetcsv($h, 2000, ';')) !== false) {
        // Limpeza rigorosa do número do processo
        $proc = preg_replace('/[^0-9]/', '', trim($d[0] ?? ''));
        // Captura o procurador (Coluna J / índice 9)
        $nome = trim($d[9] ?? ''); 
        
        if ($proc !== '' && $nome !== '') {
            // Guardamos o processo como chave para busca instantânea
            $lista[$proc][] = $nome;
        }
    }
    fclose($h);
    return $lista;
}

/* ===============================
   LER ARQUIVO (MAPEAMENTO DINÂMICO)
================================ */
function lerArquivo($arquivoTmp) {
    $spreadsheet = IOFactory::load($arquivoTmp);
    $sheet = $spreadsheet->getActiveSheet();
    $rows = $sheet->toArray(null, true, true, true);
    
    $header = array_shift($rows);
    $mapping = [];

    foreach ($header as $col => $name) {
        $name = mb_strtolower(trim($name), 'UTF-8');
        
        if (strpos($name, 'processo') !== false) $mapping['processo'] = $col;
        if (strpos($name, 'tipo') !== false && strpos($name, 'comunica') !== false) $mapping['tipo'] = $col;
        if (strpos($name, 'data') !== false && strpos($name, 'comunica') !== false) $mapping['dataCom'] = $col;
        if (strpos($name, 'final') !== false && strpos($name, 'ciencia') !== false) $mapping['dataFimCiencia'] = $col;
        if ($name === 'prazo') $mapping['prazo'] = $col;
        if (strpos($name, 'tipo') !== false && strpos($name, 'prazo') !== false) $mapping['tipoPrazo'] = $col;
        
        // Prioriza 'dataCiente' (data visualizado) ou 'data da ciência'
        if (strpos($name, 'dataciente') !== false) {
            $mapping['dataCiencia'] = $col;
        } elseif (!isset($mapping['dataCiencia']) && strpos($name, 'data') !== false && strpos($name, 'ciência') !== false) {
            $mapping['dataCiencia'] = $col;
        }
        
        if (strpos($name, 'automatica') !== false || strpos($name, 'automática') !== false) $mapping['cienciaAuto'] = $col;
    }

    $dados = [];
    foreach ($rows as $row) {
        if (empty($row[$mapping['processo'] ?? 'A'])) continue;

        $dados[] = [
            'processo'       => $row[$mapping['processo'] ?? 'A'] ?? '',
            'tipo'           => $row[$mapping['tipo'] ?? 'B'] ?? '',
            'dataCom'        => $row[$mapping['dataCom'] ?? 'C'] ?? '',
            'dataFimCiencia' => $row[$mapping['dataFimCiencia'] ?? 'D'] ?? '',
            'prazo'          => $row[$mapping['prazo'] ?? 'E'] ?? '',
            'tipoPrazo'      => $row[$mapping['tipoPrazo'] ?? 'F'] ?? '',
            'dataCiencia'    => $row[$mapping['dataCiencia'] ?? 'G'] ?? '',
            'cienciaAuto'    => $row[$mapping['cienciaAuto'] ?? 'H'] ?? ''
        ];
    }
    return $dados;
}

/* ===============================
   PROCESSAMENTO
================================ */
if (isset($_POST['processar']) && isset($_FILES['arquivo'])) {
    $acervo = carregarAcervo();
    try {
        $linhas = lerArquivo($_FILES['arquivo']['tmp_name']);
    } catch (Exception $e) { die('Erro: ' . $e->getMessage()); }

    $saida = new Spreadsheet();
    
    // DEFINIR FONTE ARIAL PARA TODO O ARQUIVO
    $saida->getDefaultStyle()->getFont()->setName('Arial')->setSize(10);

    $res = $saida->getActiveSheet();
    $res->setTitle('Resultado');

    $res->fromArray([
        'Processo', 'Tipo de Comunicação', 'Data da Comunicação', 'Data Final p/ Ciência', 
        'Prazo', 'Tipo de Prazo', 'Data da Ciência (Visualizado)', 'Ciência Automática', 'Procurador'
    ], null, 'A1');

    $nao = $saida->createSheet();
    $nao->setTitle('Não Localizados');
    $nao->fromArray(['Processo', 'Tipo', 'Data'], null, 'A1');

    $lr = 2; $ln = 2;

    foreach ($linhas as $l) {
        $procLimpo = preg_replace('/[^0-9]/', '', (string)$l['processo']);

        $procurador = 'PROCURADOR NÃO LOCALIZADO';
        if (isset($acervo[$procLimpo])) {
            $procurador = implode(' / ', array_unique($acervo[$procLimpo]));
        }

        $formatData = function($val) {
            if (empty($val) || $val == '1899-12-31') return '';
            try {
                if (is_numeric($val) && $val > 1000) {
                    return Date::excelToDateTimeObject($val)->format('d/m/Y');
                }
                $dt = new DateTime(str_replace('/', '-', $val));
                return $dt->format('d/m/Y');
            } catch (Exception $e) { return $val; }
        };

        // Escrita das colunas
        $res->setCellValueExplicit('A'.$lr, $l['processo'], DataType::TYPE_STRING);
        $res->setCellValue('B'.$lr, $l['tipo']);
        $res->setCellValue('C'.$lr, $formatData($l['dataCom']));
        $res->setCellValue('D'.$lr, $formatData($l['dataFimCiencia']));
        
        $pVal = preg_replace('/[^0-9]/', '', (string)$l['prazo']);
        $res->setCellValueExplicit('E'.$lr, ($pVal !== '' ? (int)$pVal : 0), DataType::TYPE_NUMERIC);
        
        $res->setCellValue('F'.$lr, strtoupper($l['tipoPrazo']));
        $res->setCellValue('G'.$lr, $formatData($l['dataCiencia']));
        $res->setCellValue('H'.$lr, $l['cienciaAuto']);
        $res->setCellValue('I'.$lr, $procurador);

        if ($procurador === 'PROCURADOR NÃO LOCALIZADO') {
            $res->getStyle('I'.$lr)->getFont()->getColor()->setARGB('FFFF0000');
            $nao->setCellValueExplicit('A'.$ln, $l['processo'], DataType::TYPE_STRING);
            $nao->setCellValue('B'.$ln, $l['tipo']);
            $nao->setCellValue('C'.$ln, $formatData($l['dataCom']));
            $ln++;
        }
        $lr++;
    }

    foreach (range('A','I') as $col) { 
        $res->getColumnDimension($col)->setAutoSize(true); 
    }

    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment; filename="RESULTADO_PGM.xlsx"');
    (new Xlsx($saida))->save('php://output');
    exit;
}
?>
<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="UTF-8">
<title>Processador DJE</title>
<style>
body {
    font-family: 'Segoe UI', sans-serif;
    background: #f4f7f6;
    display: flex;
    justify-content: center;
    align-items: center;
    height: 100vh;
}
.box {
    background: #fff;
    padding: 40px;
    border-radius: 12px;
    box-shadow: 0 8px 30px rgba(0,0,0,.1);
    text-align: center;
}
input { width: 100%; margin-bottom: 20px; }
button {
    background: #27ae60;
    color: #fff;
    border: none;
    padding: 12px;
    border-radius: 6px;
    font-weight: bold;
    cursor: pointer;
    width: 100%;
}
</style>
</head>
<body>
<div class="box">
    <h2>Processador DJE / PGM</h2>
    <form method="post" enctype="multipart/form-data">
        <input type="file" name="arquivo" accept=".xlsx,.xls,.csv" required>
        <button type="submit" name="processar">Processar e Gerar Resultado</button>
    </form>
</div>
</body>
</html>