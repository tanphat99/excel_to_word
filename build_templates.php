<?php
/**
 * Sinh 2 file template có ${placeholder} từ 2 file Word mẫu.
 *
 *   BienBanGiaoNhanHangHoa.docx  ->  template_BBGN_word.docx
 *   DonDatHang.docx              ->  template_DDH.docx
 *
 * Word cắt vụn một câu thành nhiều <w:r>, khiến TemplateProcessor không nhận ra
 * placeholder. Script gộp các run liền kề cùng định dạng lại rồi mới chèn ${...}.
 *
 * Chạy lại mỗi khi sửa layout của 2 file mẫu:  php build_templates.php
 */

const W_NS   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';

// 👉 Bảng thay thế cho Biên Bản Giao Nhận: text trong file mẫu => text có placeholder
$bbgnMap = [
    'Hà Nội, Ngày 23 tháng 12 năm 2025'
        => 'Hà Nội, Ngày ${BBGN_Ngay} tháng ${BBGN_Thang} năm ${BBGN_Nam}',
    'Căn cứ vào hợp đồng nguyên tắc số: 012025/HĐNT PT- PCCC Trần Gia  ký ngày 10 tháng 12 năm 2025, giữa Công ty TNHH PCCC Trần Gia và Công ty TNHH sản xuất và thương mại dịch vụ Phú Thư;'
        => 'Căn cứ vào hợp đồng nguyên tắc số: ${HDNT_So} ký ngày ${HDNT_Ngay} tháng ${HDNT_Thang} năm ${HDNT_Nam}, giữa ${BenA_Ten} và Công ty TNHH sản xuất và thương mại dịch vụ Phú Thư;',
    '1312/2025/ ĐĐH ngày 1312/2025;'
        => '${DDH_So} ngày ${DDH_NgayGon};',
    'Hôm nay, Ngày 23 tháng 12 năm 2025.Tại Công ty TNHH SẢN XUẤT VÀ THƯƠNG MẠI DỊCH VỤ PHÚ THƯ '
        => 'Hôm nay, Ngày ${BBGN_Ngay} tháng ${BBGN_Thang} năm ${BBGN_Nam}. Tại Công ty TNHH SẢN XUẤT VÀ THƯƠNG MẠI DỊCH VỤ PHÚ THƯ ',
    'CÔNG TY TNHH PCCC TRẦN GIA' => '${BenA_Ten}',
    'Bà Trần Thị THANH HUYỀN'    => '${BenA_DaiDien}',
    'Giám đốc '                  => '${BenA_ChucVu}',
    'Địa chỉ: Số nhà 8, Ngách 7/24, Ngõ 311, Xóm Tiền Phong, Xã An Khánh, TP Hà Nội, Việt Nam'
        => 'Địa chỉ: ${BenA_DiaChi}',
    'Mã số thuế: 0111167893' => 'Mã số thuế: ${BenA_MaSoThue}',
    'Số điện thoại: '        => 'Số điện thoại: ${BenA_DienThoai}',
    'Số tài khoản: 2299686868 tại ngân hàng TMCP Sài Gòn- Hà Nội – Chi nhánh Hà thành'
        => 'Số tài khoản: ${BenA_TaiKhoan}',
    '1312/2025/ ĐĐH ngày 1312/2025' => '${DDH_So} ngày ${DDH_NgayGon}',
    'Thép cuộn các loại'            => '${TenHangHoa}',
    'kg'                            => '${DonViTinh}',
    '28.280'                        => '${SoLuong}',
    '10.300'                        => '${DonGia}',
    '291.284.000'                   => '${ThanhTien}',
    'Thuế GTGT 10%'                 => 'Thuế GTGT ${ThueSuat}',
    '29.128.400'                    => '${TienThueGTGT}',
    '320.412.400 '                  => '${TongTienThanhToan}',
    'Ba trăm hai mươi triệu bốn trăm mười hai nghìn bốn trăm đồng chẵn.'
        => '${TongTienBangChu}.',
    // 3 lần xuất hiện: chức vụ bên B, ô ký ĐẠI DIỆN BÊN A, ô ký ĐẠI DIỆN BÊN B
    'Giám đốc' => [null, '${BenA_ChucVu}', null],
];

// 👉 Bảng thay thế cho Đơn Đặt Hàng
$ddhMap = [
    'CÔNG TY TNHH PCCC TRẦN GIA' => '${BenA_Ten}',
    'Số: 1312/2025/ ĐĐH'         => 'Số: ${DDH_So}',
    '- Căn cứ vào hợp đồng nguyên tắc số: 012025/HĐNT PT- PCCC Trần Gia ký ngày 10 tháng 12 năm 2025, giữa Công ty TNHH PCCC Trần Gia và '
        => '- Căn cứ vào hợp đồng nguyên tắc số: ${HDNT_So} ký ngày ${HDNT_Ngay} tháng ${HDNT_Thang} năm ${HDNT_Nam}, giữa ${BenA_Ten} và ',
    'Bà Trần Thị Thanh Huyền' => '${BenA_DaiDien}',
    'Giám đốc '               => '${BenA_ChucVu}',
    'Địa chỉ: Số nhà 8, Ngách 7/24, Ngõ 311, Xóm Tiền Phong, Xã An Khánh, TP Hà Nội, Việt Nam'
        => 'Địa chỉ: ${BenA_DiaChi}',
    'Mã số thuế: 0111167893' => 'Mã số thuế: ${BenA_MaSoThue}',
    'Số điện thoại: '        => 'Số điện thoại: ${BenA_DienThoai}',
    'Số tài khoản: 2299686868 tại ngân hàng TMCP Sài Gòn- Hà Nội – Chi nhánh Hà thành'
        => 'Số tài khoản: ${BenA_TaiKhoan}',
    'Hôm nay, ngày 13 tháng 12 năm 2025 Công ty TNHH PCCC Trần Gia có nhu cầu đặt hàng tại Quý Công ty như sau:'
        => 'Hôm nay, ngày ${DDH_Ngay} tháng ${DDH_Thang} năm ${DDH_Nam} ${BenA_Ten} có nhu cầu đặt hàng tại Quý Công ty như sau:',
    'Thép cuộn các loại' => '${TenHangHoa}',
    'Kg'                 => '${DonViTinh}',
    '28.280'             => '${SoLuong}',
    '10.300'             => '${DonGia}',
];

buildTemplate(__DIR__ . '/BienBanGiaoNhanHangHoa.docx', __DIR__ . '/template_BBGN_word.docx', $bbgnMap);
buildTemplate(__DIR__ . '/DonDatHang.docx',             __DIR__ . '/template_DDH.docx',       $ddhMap);

function buildTemplate($source, $target, array $map) {
    $doc = loadDocumentXml($source);
    normalizeRuns($doc);

    $seen     = [];
    $replaced = array_fill_keys(array_keys($map), 0);

    foreach (textNodes($doc) as $textNode) {
        $text = canonical($textNode->textContent);
        if (!isset($map[$text])) continue;

        // Với map dạng mảng: phần tử thứ n ứng với lần xuất hiện thứ n, null = giữ nguyên
        $occurrence = $seen[$text] = ($seen[$text] ?? -1) + 1;
        $to = is_array($map[$text]) ? ($map[$text][$occurrence] ?? null) : $map[$text];
        if ($to === null) continue;

        setText($textNode, $to);
        $replaced[$text]++;
    }

    $missing = [];
    foreach ($map as $text => $to) {
        $expected = is_array($to) ? count(array_filter($to, function ($v) { return $v !== null; })) : 1;
        if ($replaced[$text] < $expected) {
            $missing[] = $text . "  (thay được {$replaced[$text]}/{$expected})";
        }
    }
    if ($missing) {
        throw new RuntimeException(basename($source) . " — không khớp đoạn:\n  - " . implode("\n  - ", $missing));
    }

    copy($source, $target);
    $zip = new ZipArchive();
    $zip->open($target);
    $zip->addFromString('word/document.xml', $doc->saveXML());
    $zip->close();

    echo basename($target) . ' ✔ ' . array_sum($replaced) . " đoạn đã thay\n";
}

/** Word hay chèn non-breaking space (U+00A0) lẫn với space thường — quy về một loại để so khớp */
function canonical($text) {
    return str_replace("\xC2\xA0", ' ', $text);
}

function loadDocumentXml($docxPath) {
    $zip = new ZipArchive();
    if ($zip->open($docxPath) !== true) throw new RuntimeException("Không mở được $docxPath");
    $xml = $zip->getFromName('word/document.xml');
    $zip->close();

    $doc = new DOMDocument();
    $doc->preserveWhiteSpace = true;
    $doc->loadXML($xml);
    return $doc;
}

/** Gộp các <w:r> liền kề có cùng định dạng, để một câu nằm trọn trong một run */
function normalizeRuns(DOMDocument $doc) {
    $xpath = new DOMXPath($doc);
    $xpath->registerNamespace('w', W_NS);

    // Các node chỉ phục vụ soát chính tả / phân trang, chen giữa các run và chặn việc gộp
    $noise = ['w:proofErr', 'w:bookmarkStart', 'w:bookmarkEnd', 'w:lastRenderedPageBreak',
              'w:rPr/w:lang', 'w:rPr/w:noProof'];
    foreach ($noise as $tag) {
        foreach (iterator_to_array($xpath->query('//' . $tag)) as $node) {
            $node->parentNode->removeChild($node);
        }
    }

    foreach ($xpath->query('//w:p') as $paragraph) {
        $prev = null;
        foreach (iterator_to_array($paragraph->childNodes) as $child) {
            if (!($child instanceof DOMElement) || $child->localName !== 'r' || !isPlainTextRun($child)) {
                $prev = null;
                continue;
            }
            if ($prev && runProperties($prev) === runProperties($child)) {
                setText(textNode($prev), $prev->textContent . $child->textContent);
                $paragraph->removeChild($child);
            } else {
                $prev = $child;
            }
        }
    }
}

/** Chỉ gộp run thuần text — run chứa tab, ngắt dòng hay ảnh thì giữ nguyên */
function isPlainTextRun(DOMElement $run) {
    $textCount = 0;
    foreach ($run->childNodes as $child) {
        if (!($child instanceof DOMElement)) continue;
        if ($child->localName === 'rPr') continue;
        if ($child->localName === 't') { $textCount++; continue; }
        return false;
    }
    return $textCount === 1;
}

function runProperties(DOMElement $run) {
    foreach ($run->childNodes as $child) {
        if ($child instanceof DOMElement && $child->localName === 'rPr') {
            return $run->ownerDocument->saveXML($child);
        }
    }
    return '';
}

function textNode(DOMElement $run) {
    foreach ($run->childNodes as $child) {
        if ($child instanceof DOMElement && $child->localName === 't') return $child;
    }
    return null;
}

function textNodes(DOMDocument $doc) {
    $xpath = new DOMXPath($doc);
    $xpath->registerNamespace('w', W_NS);
    return iterator_to_array($xpath->query('//w:t'));
}

function setText(DOMElement $textNode, $value) {
    $textNode->nodeValue = '';
    $textNode->appendChild($textNode->ownerDocument->createTextNode($value));
    $textNode->setAttributeNS(XML_NS, 'xml:space', 'preserve');
}
