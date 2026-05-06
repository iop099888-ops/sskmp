<?php
// Подключаем установленные библиотеки
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate,
    PhpOffice\PhpSpreadsheet\Style\Alignment,
    PhpOffice\PhpSpreadsheet\Style\Border;

//
// --- 1. КОНФИГУРАЦИЯ ---
const NC_BASE_URL = 'https://cloud.sskuban.ru/remote.php/dav/files/egorov.ee/';
const NC_USERNAME = 'egorov.ee';
const NC_PASSWORD = 'Az07031984';
const SHARE_URL = 'https://cloud.sskuban.ru/s/JRqdp5gyykssLtD';
const SHARE_PASSWORD = 'JRqdp5gyykssLtD';
const SOURCE_FILE = 'Ежемесячный и квартальный Шаблон МП Q2.xlsx';
const TARGET_SHEET_NAME = 'Справочник Источники';

// --- 2. ОБРАБОТЧИК ---
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['source_file'])) {
    try {
        $uploadedFile = $_FILES['source_file']['tmp_name'];
        if (!file_exists($uploadedFile)) {
            throw new Exception('Ошибка загрузки файла.');
        }
        
        $newData = parseUploadedFile($uploadedFile);
        $targetFilePath = findTargetFile();
        $tempFile = downloadFile($targetFilePath);
        updateFile($tempFile, $newData);
        uploadFile($tempFile, $targetFilePath);
        
        echo json_encode(['success' => true, 'message' => 'Справочник успешно обновлён!']);
    } catch (Exception $e) {
        error_log("SCSSync Error: " . $e->getMessage());
        echo json_encode(['success' => false, 'message' => 'Ошибка: ' . $e->getMessage()]);
    }
    exit;
}

// --- 3. HTML Интерфейс ---
?>
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Обновление справочника маркетинговых источников Nextcloud</title>
    <style>
        body { font-family: system-ui, 'Segoe UI', Roboto, 'Helvetica Neue', sans-serif; background: #f0f4f8; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; padding: 20px; }
        .card { background: white; border-radius: 24px; box-shadow: 0 12px 30px rgba(0,0,0,0.1); padding: 32px; max-width: 520px; width: 100%; text-align: center; transition: all 0.2s ease; border: 1px solid #e9edf2; }
        h1 { font-size: 1.8rem; font-weight: 600; margin-bottom: 0.5rem; color: #1a2c3e; letter-spacing: -0.3px; }
        .sub { color: #5c6f87; border-bottom: 1px solid #eef2f6; padding-bottom: 20px; margin-bottom: 24px; font-size: 0.95rem; }
        .upload-area { border: 2px dashed #cddae9; border-radius: 20px; padding: 32px 20px; background: #fafcff; margin-bottom: 24px; transition: border-color 0.2s; cursor: pointer; }
        .upload-area:hover { border-color: #2c7da0; background: #f6fafe; }
        input[type="file"] { display: none; }
        label.btn-upload { background: #eef2fa; padding: 10px 24px; border-radius: 60px; font-weight: 500; color: #1f6392; cursor: pointer; display: inline-block; transition: all 0.2s; font-size: 0.9rem; }
        label.btn-upload:hover { background: #e2e9f3; transform: translateY(-1px); }
        .file-name { margin-top: 16px; font-size: 0.85rem; color: #2c3e66; background: #f5f7fb; display: inline-block; padding: 6px 16px; border-radius: 60px; }
        button { background: #1f6392; border: none; padding: 12px 28px; border-radius: 40px; color: white; font-weight: 600; font-size: 1rem; margin: 16px 0 24px; cursor: pointer; transition: all 0.2s; box-shadow: 0 2px 6px rgba(0,0,0,0.05); width: auto; min-width: 180px; }
        button:hover { background: #0f4b70; transform: translateY(-2px); box-shadow: 0 8px 18px rgba(31,99,146,0.2); }
        .message { padding: 14px 20px; border-radius: 100px; font-size: 0.9rem; background: #f0f7fb; margin-top: 16px; color: #1f6392; font-weight: 500; transition: all 0.2s; }
        .error { background: #fee9e6; color: #b13e3e; }
        .success { background: #e1f7e8; color: #1e6f3f; }
        footer { font-size: 0.7rem; color: #8e9eae; margin-top: 24px; border-top: 1px solid #edf2f7; padding-top: 20px; }
        hr { margin: 16px 0; border: none; height: 1px; background: linear-gradient(to right, #e2e8f0, transparent); }
    </style>
</head>
<body>
<div class="card">
    <h1>📊 Обновить справочник</h1>
    <div class="sub">синхронизация · новый лист «Справочник Источники»</div>
    
    <div class="upload-area" id="dropzone">
        <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="#5c6f87" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round" style="margin-bottom: 12px;"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
        <p style="color:#2d4259;">Перетащите файл сюда или</p>
        <label for="fileInput" class="btn-upload">📂 Выберите файл (XLSX/XLS)</label>
        <input type="file" id="fileInput" accept=".xlsx, .xls">
        <div id="fileName" class="file-name" style="display:none;"></div>
    </div>
    
    <button id="syncBtn">🚀 Запустить обновление</button>
    <div id="resultMsg" class="message">Готов к работе. Выберите файл справочника.</div>
    <footer>После запуска заменится лист <strong>«Справочник Источники»</strong> в файле <strong><?= htmlspecialchars(SOURCE_FILE) ?></strong><br>Nextcloud WebDAV · защищённое соединение</footer>
</div>

<script>
    const fileInput = document.getElementById('fileInput');
    const fileNameDiv = document.getElementById('fileName');
    const syncBtn = document.getElementById('syncBtn');
    const resultMsg = document.getElementById('resultMsg');
    let selectedFile = null;

    fileInput.addEventListener('change', function(e) {
        if (e.target.files.length) {
            selectedFile = e.target.files[0];
            fileNameDiv.textContent = `📎 ${selectedFile.name}`;
            fileNameDiv.style.display = 'inline-block';
            resultMsg.textContent = 'Файл выбран, нажмите «Запустить обновление».';
            resultMsg.classList.remove('error', 'success');
            resultMsg.classList.add('message');
        }
    });

    const dropzone = document.getElementById('dropzone');
    dropzone.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.style.borderColor = '#2c7da0'; dropzone.style.background = '#f4f9fe'; });
    dropzone.addEventListener('dragleave', (e) => { e.preventDefault(); dropzone.style.borderColor = '#cddae9'; dropzone.style.background = '#fafcff'; });
    dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.style.borderColor = '#cddae9';
        dropzone.style.background = '#fafcff';
        if (e.dataTransfer.files.length) {
            selectedFile = e.dataTransfer.files[0];
            fileInput.files = e.dataTransfer.files;
            fileNameDiv.textContent = `📎 ${selectedFile.name}`;
            fileNameDiv.style.display = 'inline-block';
            resultMsg.textContent = 'Файл готов, нажмите кнопку.';
            resultMsg.classList.remove('error', 'success');
        }
    });

    syncBtn.addEventListener('click', async () => {
        if (!selectedFile) {
            resultMsg.textContent = '❌ Пожалуйста, сначала выберите XLSX/XLS файл.';
            resultMsg.classList.add('error');
            resultMsg.classList.remove('success', 'message');
            return;
        }
        resultMsg.textContent = '🔄 Синхронизация... обновляем справочник Nextcloud.';
        resultMsg.classList.remove('error', 'success');
        const formData = new FormData();
        formData.append('source_file', selectedFile);
        try {
            const response = await fetch(window.location.href, { method: 'POST', body: formData });
            const data = await response.json();
            if (data.success) {
                resultMsg.classList.add('success');
                resultMsg.textContent = '✅ ' + data.message;
            } else {
                throw new Error(data.message || 'Ошибка обработки.');
            }
        } catch (err) {
            resultMsg.classList.add('error');
            resultMsg.textContent = '⚠️ ' + err.message;
        }
    });
</script>

<?php
// ========== ФУНКЦИИ ОБРАБОТКИ ==========
function parseUploadedFile($filePath): array {
    $spreadsheet = IOFactory::load($filePath);
    $worksheet = $spreadsheet->getActiveSheet();
    $rows = $worksheet->toArray();
    if (empty($rows)) {
        throw new Exception('Загруженный файл не содержит данных.');
    }
    return $rows;
}

function findTargetFile(): string {
    $shareToken = parse_url(SHARE_URL, PHP_URL_PATH);
    $shareToken = ltrim($shareToken, '/s/');
    $shareToken = explode('?', $shareToken)[0];
    
    $publicDavUrl = 'https://cloud.sskuban.ru/public.php/webdav/';
    $ch = curl_init();
    curl_setopt_array($ch, [
        CURLOPT_URL => $publicDavUrl,
        CURLOPT_USERPWD => $shareToken . ':' . SHARE_PASSWORD,
        CURLOPT_CUSTOMREQUEST => 'PROPFIND',
        CURLOPT_HTTPHEADER => ['Depth: 1', 'Content-Type: application/xml'],
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_SSL_VERIFYPEER => false,
        CURLOPT_SSL_VERIFYHOST => false
    ]);
    
    $response = curl_exec($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);
    
    if ($httpCode !== 207 && $httpCode !== 200) {
        throw new Exception("Не удалось получить содержимое общей папки. HTTP код: $httpCode");
    }
    
    $xml = simplexml_load_string($response);
    $namespaces = $xml->getNamespaces(true);
    $hrefs = $xml->children($namespaces['d'] ?? 'DAV:')->response->children($namespaces['d'] ?? 'DAV:')->href;
    
    if (!$hrefs) {
        foreach ($xml->children('DAV:')->response as $response) {
            $href = (string)$response->children('DAV:')->href;
            if (basename($href) === SOURCE_FILE) {
                $rawPath = rawurldecode($href);
                if (strpos($rawPath, '/public.php/webdav/') === 0) {
                    $relativePath = substr($rawPath, strlen('/public.php/webdav/'));
                } else {
                    $relativePath = ltrim($rawPath, '/');
                }
                return $relativePath;
            }
        }
        throw new Exception("Файл '".SOURCE_FILE."' не найден в общей папке.");
    }
    
    $relativePath = null;
    foreach ($hrefs as $href) {
        $hrefDecoded = rawurldecode((string)$href);
        if (basename($hrefDecoded) === SOURCE_FILE) {
            $cleanPath = ltrim($hrefDecoded, '/');
            if (strpos($cleanPath, 'public.php/webdav/') === 0) {
                $relativePath = substr($cleanPath, strlen('public.php/webdav/'));
            } else {
                $relativePath = $cleanPath;
            }
            break;
        }
    }
    if (!$relativePath) {
        throw new Exception("Файл '".SOURCE_FILE."' не найден в общей папке.");
    }
    return $relativePath;
}

function downloadFile($remotePath): string {
    $fileUrl = NC_BASE_URL . ltrim($remotePath, '/');
    $ch = curl_init();
    curl_setopt_array($ch, [
        CURLOPT_URL => $fileUrl,
        CURLOPT_USERPWD => NC_USERNAME . ':' . NC_PASSWORD,
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_SSL_VERIFYPEER => false,
        CURLOPT_SSL_VERIFYHOST => false,
        CURLOPT_FOLLOWLOCATION => true
    ]);
    $data = curl_exec($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);
    
    if ($httpCode !== 200) {
        throw new Exception("Ошибка скачивания файла. Код: $httpCode");
    }
    $tempFile = tempnam(sys_get_temp_dir(), 'nc_update_');
    file_put_contents($tempFile, $data);
    return $tempFile;
}

function updateFile($tempFile, array $newData): void {
    $spreadsheet = IOFactory::load($tempFile);
    
    // Если лист существует — удалим
    if ($spreadsheet->sheetNameExists(TARGET_SHEET_NAME)) {
        $sheetIndex = $spreadsheet->getIndex(
            $spreadsheet->getSheetByName(TARGET_SHEET_NAME)
        );
        $spreadsheet->removeSheetByIndex($sheetIndex);
    }
    
    // Создаём новый лист
    $newSheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet(
        $spreadsheet,
        TARGET_SHEET_NAME
    );
    $spreadsheet->addSheet($newSheet);
    
    // Заполняем лист данными
    $highestColumnLetter = Coordinate::stringFromColumnIndex(count($newData[0] ?? [1]));
    foreach ($newData as $rowIndex => $rowData) {
        $row = $rowIndex + 1;
        $colIndex = 0;
        foreach ($rowData as $cellValue) {
            $colIndex++;
            $columnLetter = Coordinate::stringFromColumnIndex($colIndex);
            $cell = $newSheet->getCell($columnLetter . $row);
            $cell->setValueExplicit(
                $cellValue ?? '',
                \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING
            );
        }
    }
    
    // Автоширина для всех колонок
    foreach (range('A', $highestColumnLetter) as $columnID) {
        $newSheet->getColumnDimension($columnID)->setAutoSize(true);
    }
    
    // Стилизация шапки
    $headerRange = 'A1:' . $highestColumnLetter . '1';
    $newSheet->getStyle($headerRange)->getFont()->setBold(true);
    $newSheet->getStyle($headerRange)->getAlignment()
        ->setHorizontal(Alignment::HORIZONTAL_CENTER);
    $newSheet->getStyle($headerRange)->getBorders()->getBottom()
        ->setBorderStyle(Border::BORDER_THIN);
    
    $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
    $writer->save($tempFile);
}

function uploadFile($localPath, $remotePath): void {
    $content = file_get_contents($localPath);
    if ($content === false) {
        throw new Exception('Не удалось прочитать обновлённый файл.');
    }
    $uploadUrl = NC_BASE_URL . ltrim($remotePath, '/');
    $ch = curl_init();
    curl_setopt_array($ch, [
        CURLOPT_URL => $uploadUrl,
        CURLOPT_USERPWD => NC_USERNAME . ':' . NC_PASSWORD,
        CURLOPT_CUSTOMREQUEST => 'PUT',
        CURLOPT_POSTFIELDS => $content,
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_SSL_VERIFYPEER => false,
        CURLOPT_SSL_VERIFYHOST => false,
        CURLOPT_HTTPHEADER => ['Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']
    ]);
    curl_exec($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);
    
    if ($httpCode !== 201 && $httpCode !== 204 && $httpCode !== 200) {
        throw new Exception("Ошибка загрузки файла. HTTP код: $httpCode");
    }
}
?>
