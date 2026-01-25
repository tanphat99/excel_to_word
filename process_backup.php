<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpWord\TemplateProcessor;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory as WordIOFactory;

if (!isset($_FILES['excel_file'])) {
    echo "<div style='
        font-family: Arial, sans-serif;
        background-color: #fff3cd;
        color: #856404;
        padding: 20px;
        margin: 50px auto;
        border: 1px solid #ffeeba;
        border-radius: 8px;
        width: fit-content;
        text-align: center;
    '>
        <p><strong>Lỗi:</strong> Không có file nào được upload.</p>
        <a href='index.php' style='
            display: inline-block;
            margin-top: 10px;
            padding: 10px 20px;
            background-color: #007bff;
            color: white;
            text-decoration: none;
            border-radius: 6px;
            font-weight: bold;
        '>Quay lại</a>
    </div>";
    exit;
}
$docType = $_POST['doc_type'] ?? 'HĐMB';

$tmpPath = $_FILES['excel_file']['tmp_name'];
$spreadsheet = IOFactory::load($tmpPath);
$sheet = $spreadsheet->getActiveSheet();
$data = $sheet->toArray(null, true, true, true);

$headers = array_shift($data); // Lấy tiêu đề cột

$outputDir = __DIR__ . '/output/';
if (!is_dir($outputDir)) mkdir($outputDir, 0777, true);

$index = 1;
$currentYear = date('Y'); // Lấy năm hiện tại
$placeholders = [
    ['col' => 'A',  'index' => 1,  'name' => 'STT'],
    ['col' => 'B',  'index' => 2,  'name' => 'MauSo'],
    ['col' => 'C',  'index' => 3,  'name' => 'KyHieu'],
    ['col' => 'D',  'index' => 4,  'name' => 'SoHoaDon'],
    ['col' => 'E',  'index' => 5,  'name' => 'NgayThangNamHD'],
    ['col' => 'F',  'index' => 6,  'name' => 'HoTenNguoiMua'],
    ['col' => 'G',  'index' => 7,  'name' => 'CCCD'],
    ['col' => 'H',  'index' => 8,  'name' => 'MaHoChieu'],
    ['col' => 'I',  'index' => 9,  'name' => 'MaDVQHNS'],
    ['col' => 'J',  'index' => 10, 'name' => 'TenDonVi'],
    ['col' => 'K',  'index' => 11, 'name' => 'DiaChi'],
    ['col' => 'L',  'index' => 12, 'name' => 'MaSoThue'],
    ['col' => 'M',  'index' => 13, 'name' => 'TenHangHoaDichVu'],
    ['col' => 'N',  'index' => 14, 'name' => 'SoLuong'],
    ['col' => 'O',  'index' => 15, 'name' => 'DonGia'],
    ['col' => 'P',  'index' => 16, 'name' => 'ThanhTien'],
    ['col' => 'Q',  'index' => 17, 'name' => 'ThueSuat'],
    ['col' => 'R',  'index' => 18, 'name' => 'TienThueGTGT'],
    ['col' => 'S',  'index' => 19, 'name' => 'TongTienThanhToan'],
    ['col' => 'T',  'index' => 20, 'name' => 'HinhThucTT'],
    ['col' => 'U',  'index' => 21, 'name' => 'MTC'],
    ['col' => 'V',  'index' => 22, 'name' => 'MaCuaCQT'],
    ['col' => 'W',  'index' => 23, 'name' => 'TrangThaiHoaDon'],
    ['col' => 'X',  'index' => 24, 'name' => 'GhiChu'],
    ['col' => 'Y',  'index' => 25, 'name' => 'TienTe'],
    ['col' => 'Z',  'index' => 26, 'name' => 'TyGia'],
];

$companies = [
    '0108947484' => [
        'ten'      => 'CÔNG TY CỔ PHẦN THIẾT BỊ ĐIỆN KING',
        'code'     => 'K',
        'code_name'=> 'KING',
        'diachi'   => 'Thôn Từ Am – Xã Thanh Thùy – Huyện Thanh Oai - Thành phố Hà Nội, Việt Nam',
        'sdt'      => '0929.112.999',
        'fax'      => '',
        'daidien'  => 'Ông Nguyễn Văn Tân',
        'chucvu'   => 'Giám đốc',
        'taikhoan' => '2662688888 – Tại Ngân hàng TMCP Á Châu – Chi nhánh Thanh Trì – Hà Nội',
    ],
    '0104315188' => [
        'ten'      => 'CÔNG TY CỔ PHẦN SẢN XUẤT CƠ KHÍ VÀ THƯƠNG MẠI  THÁI DƯƠNG',
        'code'     => 'TD',
        'code_name'=> 'THAI_DUONG',
        'diachi'   => 'Thôn Từ Am – Xã Thanh Thùy – Huyện Thanh Oai - Thành phố Hà Nội, Việt Nam',
        'sdt'      => '0986.730.783',
        'fax'      => '',
        'daidien'  => 'Ông Thái Đình Vinh',
        'chucvu'   => 'Giám đốc',
        'taikhoan' => "26626888 – Tại Ngân hàng TMCP Á Châu – Chi nhánh Thanh Trì – Hà Nội.\n0101100118999 – Tại ngân hàng TMCP Quân Đội – CN Sở giao dịch 1.",
    ],
    '0111078788' => [
        'ten'      => 'CÔNG TY TNHH SẢN XUẤT CƠ KHÍ VÀ THUƠNG MẠI TUẤN KIỆT',
        'code'     => 'TK',
        'code_name'=> 'TUAN_KIET',
        'diachi'   => 'Số nhà 35, ngõ Đồng Đanh, thôn Rùa Thượng, xã Thanh Thùy, Huyện Thanh Oai,Thành Phố Hà Nội, Việt Nam.',
        'sdt'      => '',
        'fax'      => '',
        'daidien'  => 'Ông Nguyễn Văn Bách',
        'chucvu'   => 'Giám đốc',
        'taikhoan' => '',
    ],
    '0107172756' => [
        'ten'      => 'CÔNG TY TNHH SẢN XUẤT VÀ THƯƠNG MẠI CƠ KHÍ TIẾN THÀNH',
        'code'     => 'TT',
        'code_name'=> 'TIEN_THANH',
        'diachi'   => 'Thôn Rùa Hạ, Xã Thanh Thùy, Huyện Thanh Oai, Thành phố Hà Nội, Việt Nam',
        'sdt'      => '0986.132.374',
        'fax'      => '',
        'daidien'  => 'Ông Lý Văn Đấu',
        'chucvu'   => 'Giám đốc',
        'taikhoan' => '119000165843 tại Ngân hàng VietinBank',
    ],
    '0110857100' => [
        'ten'      => 'CÔNG TY TNHH THƯƠNG MAI VÀ SẢN XUẤT HOÀNG KIM VIỆT NAM',
        'code'     => 'HK',
        'code_name'=> 'HOANG_KIM',
        'diachi'   => 'Số 6, Xóm Mới, Thôn Rùa Thượng, Xã Thanh Thùy, Huyện Thanh Oai, Thành phố Hà Nội, Việt Nam',
        'sdt'      => '0337.333.886',
        'fax'      => '',
        'daidien'  => 'Lê Hoàng Thương',
        'chucvu'   => 'Giám đốc',
        'taikhoan' => '',
    ],
    '0109436210' => [
        'ten'      => 'CÔNG TY TNHH XÂY DỰNG CƠ KHÍ MINH TÂM',
        'code'     => 'MT',
        'code_name'=> 'MINH_TAM',
        'diachi'   => 'Cụm Công Nghiệp Thanh Oai, Xã Bích Hòa, Huyện Thanh Oai, Thành phố Hà Nội, Việt Nam',
        'sdt'      => '0936.133.995 - 0937.133.995',
        'fax'      => '',
        'daidien'  => 'Bà Nguyễn Thị Bích Ngọc',
        'chucvu'   => 'Giám đốc',
        'taikhoan' => '5151133995995 tại Ngân hàng Quân Đội chi nhanh Tây Hà Nội',
    ],
    '0109581881' => [
        'ten'      => 'CÔNG TY TNHH SẢN XUẤT CƠ KHÍ VÀ THƯƠNG MẠI NHẬT MINH',
        'code'     => 'NM',
        'code_name'=> 'NHAT_MINH',
        'diachi'   => 'Thôn Rùa Thượng, Xã Thanh Thùy, Huyện Thanh Oai, Thành phố Hà Nội, Việt Nam',
        'sdt'      => '',
        'fax'      => '',
        'daidien'  => 'Ông Hoàng Trường Đông',
        'chucvu'   => 'Giám đốc',
        'taikhoan' => '868388.8683888 tại Ngân hàng TMCP Quân Đội chi nhanh Tây Hà Nội',
    ],
    '0107394773' => [
        'ten'      => 'CÔNG TY TNHH THIẾT BỊ HANOIME',
        'code'     => 'HNM',
        'code_name'=> 'HANOIME',
        'diachi'   => 'Số 39B, ngõ 2 Quang Trung, Phường Quang Trung, Quận Hà Đông, Thành phố Hà Nội, Việt Nam',
        'sdt'      => '',
        'fax'      => '',
        'daidien'  => '',
        'chucvu'   => '',
        'taikhoan' => '',
    ],
    '0601194984' => [
        'ten'      => 'CÔNG TY TNHH ĐẦU TƯ VÀ SẢN XUẤT ĐẠI LINH',
        'code'     => 'DL',
        'code_name'=> 'DAI_LINH',
        'diachi'   => 'Tổ 15, Thị trấn Nam Giang, Huyện Nam Trực, Tỉnh Nam Định, Việt Nam',
        'sdt'      => '',
        'fax'      => '',
        'daidien'  => '',
        'chucvu'   => '',
        'taikhoan' => '',
    ],
    '0109691813' => [
        'ten'      => 'CÔNG TY CỔ PHẦN SẢN XUẤT VÀ THƯƠNG MẠI  ANH QUÂN',
        'code'     => 'AQ',
        'code_name'=> 'ANH_QUAN',
        'diachi'   => 'Cụm 7 Thôn Phượng Nghĩa, Xã Phụng Châu, Huyện Chương Mỹ, Thành phố Hà Nội, Việt Nam',
        'sdt'      => '',
        'fax'      => '',
        'daidien'  => '',
        'chucvu'   => '',
        'taikhoan' => '',
    ],
    '0110133651' => [
        'ten'      => 'CÔNG TY TNHH SẢN XUẤT VÀ THƯƠNG MAI KIÊN GIA',
        'code'     => 'KG',
        'code_name'=> 'KIEN_GIA',
        'diachi'   => 'LK3-19 Khu đô thị Tân Tây Đô, Xã Tân Lập, Huyện Đan Phượng, Thành phố Hà Nội, Việt Nam',
        'sdt'      => '',
        'fax'      => '',
        'daidien'  => '',
        'chucvu'   => '',
        'taikhoan' => '',
    ],
    '0500567284' => [
        'ten'      => 'CÔNG TY TNHH CƠ KHÍ THIÊN PHÚ',
        'code'     => 'TP',
        'code_name'=> 'THIEN_PHU',
        'diachi'   => 'Thôn Rùa Hạ, xã Thanh Thùy, huyện Thanh Oai, thành phố Hà Nội',
        'sdt'      => '',
        'fax'      => '',
        'daidien'  => 'Ông Nguyễn Văn Biên',
        'chucvu'   => 'Giám đốc',
        'taikhoan' => '',
    ],
   '0110985984' => [
       'ten'      => 'Công Ty TNHH Phát Triển SX & TM Khang Thịnh Phát',
       'code'     => 'KTP',
       'code_name'=> 'KHANG_THINH_PHAT',
       'diachi'   => 'Số nhà 78 Thôn Phú Đôi - Xã Văn Hoàng - Huyện Phú Xuyên -TP Hà Nội',
       'sdt'      => '0973.086.263',
       'fax'      => '',
       'daidien'  => 'Ông Nguyễn Xuân Khang',
       'chucvu'   => 'Giám đốc',
       'taikhoan' => '2208197638888 Ngân Hàng AGRIBANK Chi Nhánh Thường Tín- Số 116-Thị Trấn Thường Tín-Huyện Thường Tín-TP Hà Nội',
   ],
    '0107017038' => [
        'ten'      => 'CÔNG TY CỔ PHẦN DMEC VIỆT NAM',
        'code'     => 'DMEC',
        'code_name'=> 'DMEC',
        'diachi'   => 'Số 130 ngõ Thịnh Quang, phố Thái Thịnh, Phường Thịnh Quang, Quận Đống Đa, Thành phố Hà Nội, Việt Nam',
        'sdt'      => '',
        'fax'      => '',
        'daidien'  => 'Ông Đàm Đình Trường',
        'chucvu'   => 'Giám đốc',
        'taikhoan' => '1913 0816 3188 68 tại Ngân hàng Techcombank – Chi nhánh Trần Điền- Hà Nội',
    ],
    '0109135799' => [
        'ten'      => 'CÔNG TY TNHH SẢN XUẤT VÀ THƯƠNG MẠI THIẾT BỊ PCCC GIA PHÁT',
        'code'     => 'GP',
        'code_name'=> 'Gia_Phat',
        'diachi'   => 'Thôn Hoa Thám, Xã La Phù, Huyện Hoài Đức, Thành phố Hà Nội, Việt Nam',
        'sdt'      => '0913.325.715',
        'fax'      => '',
        'daidien'  => 'Bà Trần Thị Phuơng Hoa',
        'chucvu'   => 'Giám đốc',
        'taikhoan' => '1135797799 - Ngân hàng TMCP Sài Gòn - Hà Nội(SHB) Chi nhánh Hà Thành',
    ],
	'0109144987' => [
        'ten'      => 'CÔNG TY TNHH SẢN XUẤT KIM KHÍ HOÀNG MINH MECHANICAL',
        'code'     => 'HM',
        'code_name'=> 'Hoang_Minh_Mechanical',
        'diachi'   => 'Thôn Rùa Hạ, Xã Thanh Thùy, Huyện Thanh Oai, Thành phố Hà Nội, Việt Nam',
        'sdt'      => '',
        'fax'      => '',
        'daidien'  => 'Ông Thái Đình Kiên',
        'chucvu'   => 'Giám đốc',
        'taikhoan' => '',
    ],
	'0106029785' => [
        'ten'      => 'CÔNG TY CỔ PHẦN THƯƠNG MẠI & THIẾT BỊ TRƯỜNG AN',
        'code'     => 'TA',
        'code_name'=> 'Truong_An',
        'diachi'   => 'Số C04 khu đấu giá QSD đất Kiến Hưng - Hà Cầu, Phường Kiến Hưng, Thành phố Hà Nội, Việt Nam',
        'sdt'      => '',
        'fax'      => '',
        'daidien'  => '',
        'chucvu'   => '',
        'taikhoan' => '',
    ],
    '0101231331' => [
        'ten'      => 'CÔNG TY TNHH SẢN XUẤT VÀ THƯƠNG MẠI ANH LINH',
        'code'     => 'AL',
        'code_name'=> 'Anh_Linh',
        'diachi'   => 'Thôn Vân La - Xã Hồng Vân - Huyện Thường Tín - TP. Hà Nội',
        'sdt'      => '02438682722',
        'fax'      => '',
        'daidien'  => 'Bà Trần Thị Ngọc Lan',
        'chucvu'   => 'Phó Giám Đốc',
        'taikhoan' => '203131418 Tại Ngân hàng VPB - CN Kim Liên 2208201000166 ngân hàng NN và PTNT CN Thường Tín, Hà Nội',
    ],
    '0107723837' => [
        'ten'      => 'CÔNG TY TNHH CƠ KHÍ PHONG THÁI',
        'code'     => 'PT',
        'code_name'=> 'Phong_Thai',
        'diachi'   => 'Thôn Rùa Hạ, Xã Thanh Thùy, Huyện Thanh Oai, Thành phố Hà Nội, Việt Nam',
        'sdt'      => '',
        'fax'      => '',
        'daidien'  => 'Nguyễn Tuấn Khải',
        'chucvu'   => 'Giám Đốc',
        'taikhoan' => '110002619658 Tại Ngân hàng Công Thương Việt Nam - Chi nhánh Thành An',
    ],
    '0109066288' => [
        'ten'      => 'CÔNG TY CỔ PHẦN SẢN XUẤT VÀ XUẤT NHẬP KHẨU THƯƠNG MẠI HỢP PHÁT',
        'code'     => 'HP',
        'code_name'=> 'Hop_Phat',
        'diachi'   => 'Số nhà 4 nghách 123 ngõ 324, đường Phương Canh, Phường Xuân Phương, TP Hà Nội, Việt Nam',
        'sdt'      => '',
        'fax'      => '',
        'daidien'  => 'TRẦN TRUNG TUYẾN',
        'chucvu'   => '',
        'taikhoan' => '4520625249 tại BIDV CN Vạn Phúc - Hà Nội',
    ],
];
/**
 * @param $number
 * @return string
 */
function numberToVietnameseWords($number) {
    $formatter = new NumberFormatter("vi", NumberFormatter::SPELLOUT);
    $text = $formatter->format($number);
    $text = ucfirst($text);
    $text = str_replace(['phẩy', 'âm '], ['', ''], $text);
    return $text . ' đồng chẵn';
}


// 📌 Xử lý BBGN (Xuất file Excel mỗi dòng)
if ($docType === 'BBGN') {
    $index = 1;
    foreach ($data as $row) {
        $maSoThue = $row['L'] ?? '';
        if (!isset($companies[$maSoThue])) continue;

        $companyInfo = $companies[$maSoThue];
        $codeName = $companyInfo['code_name'];

        $templateExcel = IOFactory::load('template_output.xlsx');
        $sheet = $templateExcel->getActiveSheet();

        $originalDate = $row['E'] ?? ''; // E = cột NgayThangNamHD
        $colA6 = '';

// Nếu là số dạng Excel (timestamp)
        if (is_numeric($originalDate)) {
            $phpDate = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($originalDate);
        } else {
            $phpDate = DateTime::createFromFormat('d/m/Y', $originalDate)
                ?: DateTime::createFromFormat('d-m-Y', $originalDate);
        }

        if ($phpDate) {
            $day   = $phpDate->format('d');
            $month = $phpDate->format('m');
            $year  = $phpDate->format('Y');

            $colA6 = "Hôm nay, ngày {$day} tháng {$month} năm {$year}, chúng tôi gồm:";
        }

        $sheet->setCellValue('A6', $colA6);

        $benMua = $row['J'] ?? ''; // Cột J: TenDonVi

// Tạo RichText
        $richText = new \PhpOffice\PhpSpreadsheet\RichText\RichText();

        $boldPart = $richText->createTextRun("BÊN MUA (BÊN B): ");
        $boldPart->getFont()->setBold(true);
        $boldPart->getFont()->setName('Times New Roman');
        $boldPart->getFont()->setSize(12);

        $normalPart = $richText->createTextRun($benMua);
        $normalPart->getFont()->setName('Times New Roman');
        $normalPart->getFont()->setSize(12);

// Merge từ A10 đến I10
        $sheet->mergeCells('A10:I10');

// Gán RichText
        $sheet->setCellValue('A10', $richText);

// Bật wrap text
        $sheet->getStyle('A10')->getAlignment()->setWrapText(true);

// Căn lề trái
        $sheet->getStyle('A10')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);

// Ước tính chiều cao dòng (tùy thuộc độ dài tên công ty)
        $fullLength = mb_strlen("BÊN MUA (BÊN B): " . $benMua);
        $lines = ceil($fullLength / 50);
        $height = max(20, $lines * 20);
        $sheet->getRowDimension(10)->setRowHeight($height);


        $diaChi = $row['K'] ?? '';

        $diaChi = $row['K'] ?? '';
        $sheet->mergeCells('A11:I11');
        $sheet->setCellValue('A11', "Địa chỉ: $diaChi");
        $sheet->getStyle('A11')->getAlignment()->setWrapText(true);

// Ước lượng chiều cao: mỗi 50 ký tự ~ 15pt chiều cao
        $chars = mb_strlen($diaChi);
        $lines = ceil($chars / 50);
        $rowHeight = max(20, $lines * 15); // đảm bảo ít nhất là 20

        $sheet->getRowDimension(11)->setRowHeight($rowHeight);


        $sheet->setCellValue('B15',  $row['M']);

        $valueN = $row['N'] ?? '';
// Bỏ dấu phẩy phân cách ngàn (nếu có), sau đó ép float và cast về int để loại bỏ phần thập phân
        $cleanValue = str_replace(',', '', $valueN);
        $valueFormatted = (int) $cleanValue; // Kết quả: 7750

        $sheet->setCellValue('G15',  $valueFormatted);

        $companyDir = $outputDir . "BBGN/{$codeName}/";
        if (!is_dir($companyDir)) mkdir($companyDir, 0777, true);

        $ymd = $phpDate->format('d_m_Y');
        $fileName = "BBGN_{$ymd}_{$codeName}_{$index}.xlsx";
        $savePath = $companyDir . $fileName;
        $writer = IOFactory::createWriter($templateExcel, 'Xlsx');
        $writer->save($savePath);

        $index++;
    }

    echo "<div style='
        font-family: Arial, sans-serif;
        background-color: #d4edda;
        color: #155724;
        padding: 20px;
        margin: 50px auto;
        border: 1px solid #c3e6cb;
        border-radius: 10px;
        width: fit-content;
        text-align: center;
    '>
        <p><strong>✅ Thành công!</strong> Đã tạo file BBGN trong thư mục <code>output/BBGN/</code>.</p>
        <div style='margin-top: 15px;'>
            <a href='output/BBGN/' target='_blank' style='
                display: inline-block;
                padding: 10px 20px;
                margin-right: 10px;
                background-color: #28a745;
                color: white;
                text-decoration: none;
                border-radius: 6px;
                font-weight: bold;
            '>📂 Xem BBGN</a>
            <a href='index.php' style='
                display: inline-block;
                padding: 10px 20px;
                background-color: #007bff;
                color: white;
                text-decoration: none;
                border-radius: 6px;
                font-weight: bold;
            '>⬅️ Quay lại</a>
        </div>
    </div>";
    exit;
}

/**
 * Create Word document using simplified PhpWord approach
 */
function createWordDirectly($data, $outputPath) {
    try {
        $phpWord = new PhpWord();
        $phpWord->setDefaultFontName('Times New Roman');
        $phpWord->setDefaultFontSize(13);
        
        // Tạo hợp đồng mua bán (trang 1-2)
        $section1 = $phpWord->addSection();
        
        // Header
        $section1->addText('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', ['bold' => true, 'size' => 16], ['alignment' => 'center']);
        $section1->addText('Độc lập - Tự do - Hạnh phúc', ['bold' => true, 'size' => 14], ['alignment' => 'center']);
        $section1->addText('---- o0o ----', ['alignment' => 'center']);
        $section1->addText('HỢP ĐỒNG MUA BÁN', ['bold' => true, 'size' => 14], ['alignment' => 'center']);
        $section1->addTextBreak();
        
        // Contract Number
        $section1->addText('Số: ' . $data['NgayThangNamHD'] . '/HĐMB/PT- ' . $data['Info_TenCongTy'], ['bold' => true]);
        $section1->addTextBreak();
        
        // Legal Basis
        $section1->addText('- Căn cứ vào các Bộ luật dân sự Nước Cộng hòa Xã hội Chủ nghĩa Việt Nam');
        $section1->addText('- Căn cứ vào Luật Thương mại 2005 của Nước Cộng hòa Xã hội Chủ nghĩa Việt Nam');
        $section1->addText('- Căn cứ vào nhu cầu và khả năng cung cấp hàng hóa của các bên.');
        $section1->addTextBreak();
        
        // Date and Parties Introduction
        $section1->addText('Hôm nay, ngày    tháng    năm ' . $data['NamHienTai'] . ' chúng tôi gồm:');
        $section1->addTextBreak();
        
        // Section I: Seller
        $section1->addText('I. BÊN BÁN (BÊN A): CÔNG TY TNHH SẢN XUẤT VÀ THƯƠNG MẠI DỊCH VỤ PHÚ THƯ', ['bold' => true]);
        $section1->addText('Địa chỉ: Nhà số 1 Thôn Rùa Thượng – Xã Thanh Thùy – Huyện Thanh Oai - Thành phố Hà Nội, Việt Nam');
        $section1->addText('Điện thoại: 034.433.2918	Fax:');
        $section1->addText('Mã số thuế: 0110430598');
        $section1->addText('Đại diện: Ông Nguyễn Văn Báo           Chức vụ: Giám đốc');
        $section1->addText('Tài khoản: 117002948428 – Tại Ngân hàng Vietinbank – Chi nhánh Thành An');
        $section1->addTextBreak();
        
        // Section II: Buyer
        $section1->addText('II. BÊN MUA ( BÊN B): ' . $data['Info_TenDonViDayDu'], ['bold' => true]);
        $section1->addText('Địa chỉ: ' . $data['Info_DiaChi']);
        $section1->addText('Điện thoại: ' . $data['Info_DienThoai'] . '	Fax: ' . $data['Info_Fax']);
        $section1->addText('Mã số thuế: ' . $data['Info_MaSoThue']);
        $section1->addText('Đại diện: ' . $data['Info_DaiDien'] . '	Chức vụ: ' . $data['Info_ChucVu']);
        $section1->addText('Tài khoản: ' . $data['Info_TaiKhoan']);
        $section1->addTextBreak();
        
        // Agreement Statement
        $section1->addText('Sau khi thỏa thuận, bên B đồng ý mua, Bên A đồng ý bán hàng hóa theo các điều kiện sau đây:');
        $section1->addTextBreak();
        
        // Article 1
        $section1->addText('Điều 1: Hàng hóa, chất lượng và số lượng.', ['bold' => true]);
        $section1->addText('1.1. Hàng hóa và số lượng: Bên B đồng ý mua của bên A mặt hàng ' . $data['TenHangHoaDichVu'] . ' do bên A bán với số lượng chi tiết ở điều 2.1 của hợp đồng.');
        $section1->addText('1.2. Chất lượng: Hàng hóa mới 100%.');
        $section1->addTextBreak(3);
        
        // Article 2
        $section1->addText('Điều 2: Giá cả, phương thức thanh toán.', ['bold' => true]);
        $section1->addText('2.1. Giá cả: giá trị của hàng hóa nêu tại Điều 1 của hợp đồng này được tính theo bảng giá như sau:');
        $section1->addTextBreak();
        
        // Price Table
        $table = $section1->addTable(['borderSize' => 1]);
        
        // Header row
        $table->addRow();
        $table->addCell(1000)->addText('STT', ['bold' => true], ['alignment' => 'center']);
        $table->addCell(3000)->addText('LOẠI HÀNG', ['bold' => true], ['alignment' => 'center']);
        $table->addCell(1000)->addText('ĐVT', ['bold' => true], ['alignment' => 'center']);
        $table->addCell(1500)->addText('S.LƯỢNG', ['bold' => true], ['alignment' => 'center']);
        $table->addCell(2000)->addText('ĐƠN GIÁ', ['bold' => true], ['alignment' => 'center']);
        $table->addCell(2000)->addText('THÀNH TIỀN', ['bold' => true], ['alignment' => 'center']);
        
        // Data row
        $table->addRow();
        $table->addCell(1000)->addText('1', ['alignment' => 'center']);
        $table->addCell(3000)->addText($data['TenHangHoaDichVu']);
        $table->addCell(1000)->addText('Kg', ['alignment' => 'center']);
        $table->addCell(1500)->addText($data['SoLuong'], ['alignment' => 'right']);
        $table->addCell(2000)->addText($data['DonGia'], ['alignment' => 'right']);
        $table->addCell(2000)->addText($data['ThanhTien'], ['alignment' => 'right']);
        
        // Summary rows
        $table->addRow();
        $table->addCell(8500)->addText('Tổng cộng tiền hàng', ['bold' => true]);
        $table->addCell(2000)->addText($data['ThanhTien'], ['bold' => true, 'alignment' => 'right']);
        
        $table->addRow();
        $table->addCell(8500)->addText('Thuế VAT 10%', ['bold' => true]);
        $table->addCell(2000)->addText($data['TienThueGTGT'], ['bold' => true, 'alignment' => 'right']);
        
        $table->addRow();
        $table->addCell(8500)->addText('Tổng cộng', ['bold' => true]);
        $table->addCell(2000)->addText($data['TongTienThanhToan'], ['bold' => true, 'alignment' => 'right']);
        
        $section1->addTextBreak();
        $section1->addText('Bằng chữ: ' . $data['TongTienThanhToan_BangChu']);
        $section1->addText('Giá trị hợp đồng được phép dung sai tối đa +/-10%');
        $section1->addTextBreak();
        
        $section1->addText('2.2. Phương thức thanh toán:', ['bold' => true]);
        $section1->addText('- Hình thức thanh toán: Thanh toán trước hoặc sau không quá 60 ngày khi bên A giao đủ hàng và chứng từ cho bên B, bằng tiền mặt hoặc chuyển khoản.');
        $section1->addText('- Chứng từ thanh toán bao gồm: Hoá đơn GTGT, Biên bản giao nhận hàng hóa.');
        $section1->addTextBreak();
        
        // Article 3
        $section1->addText('Điều 3: Địa điểm và phương thức giao nhận hàng hóa', ['bold' => true]);
        $section1->addText('3.1. Địa điểm giao nhận: Tại xưởng bán bên A');
        $section1->addText('3.2. Phương thức giao nhận: Giao hàng qua cân thực tế và lập Biên bản giao nhận hàng hóa giữa hai bên.Khối lượng trên biên bản được xác định làm căn cứ thanh toán. Nếu phát hiện hàng thiếu hoặc không đúng tiêu chuẩn chất lượng thì lập biên bản tại chỗ và yêu cầu bên A xác nhận. Trường hợp hàng không đạt chất lượng như quy định tại điều 2, bên B có quyền từ chối không nhận hàng.');
        $section1->addText('3.3. Thời hạn giao hàng: Hàng giao theo yêu cầu bên mua');
        $section1->addTextBreak();
        
        // Article 4
        $section1->addText('Điều 4: Quyền và nghĩa vụ của các bên', ['bold' => true]);
        $section1->addText('4.1. Quyền và nghĩa vụ của Bên B:', ['bold' => true]);
        $section1->addText('- Kiểm tra số lượng, chất lượng hàng hóa nêu tại Điều 1 của hợp đồng này.');
        $section1->addText('- Chuẩn bị nhân lực để tiếp nhận hàng khi có thông báo giao hàng của Bên A');
        $section1->addText('- Đảm bảo việc thanh toán theo đúng giá trị và phương thức thanh toán đã nêu tại Điều 2 của Hợp đồng này.');
        $section1->addText('- Bảo quản sử dụng hàng hóa theo đúng hướng dẫn của Nhà sản xuất.');
        $section1->addText('- Có quyền khiếu nại về hàng hóa bên A giao trong thời hạn 02 ngày về số lượng, chủng loại và trong thời hạn 15 ngày về chất lượng kể từ ngày bên A giao hàng.');
        $section1->addText('4.2. Quyền và nghĩa vụ của Bên A:', ['bold' => true]);
        $section1->addText('- Yêu cầu Bên B thực hiện đúng và hai bên đã cam kết.');
        $section1->addText('- Đảm bảo cung cấp đúng và đầy đủ số lượng, chất lượng hàng hóa nêu tại Điều 1 của Hợp đồng');
        $section1->addText('- Giải quyết các khiếu nại của bên B trong thời hạn 05 ngày kể từ ngày nhận được thư khiếu nại của bên A.');
        $section1->addTextBreak();
        
        // Article 5
        $section1->addText('Điều 5: Điều khoản thi hành', ['bold' => true]);
        $section1->addText('5.1. Hai bên cam kết thực hiện đúng và đầy đủ các điều khoản đã cam kết trong Hợp đồng.');
        $section1->addText('5.2. Trong quá trình thực hiện, nếu có phát sinh những bất lợi và sự kiện mới thì hai bên cùng nhau bàn bạc, nhất trí thỏa thuận về việc sửa đổi, bổ sung Hợp đồng. Mọi sửa đổi và bổ sung của Hợp đồng chỉ có giá trị khi được hai bên thoả thuận và lập thành văn bản.');
        $section1->addText('5.3. Trường hợp có phát sinh tranh chấp thì các bên cùng tiến hành thương lượng giải quyết trên cơ sở tôn trọng quyền lợi của nhau. Nếu các bên không giải quyết được thì một trong hai bên có quyền yêu cầu Tòa án có thẩm quyền giải quyết.theo quy định của pháp luật. Phán quyết của Tòa án là căn cứ để hai bên thực hiện. Nếu bên nào sai thì hoàn toàn chịu mọi phí tổn Tòa án.');
        $section1->addText('5.5. Hợp đồng được lập thành 02 bản, mỗi bên giữ 01 bản, các bản có giá trị pháp lý như nhau.');
        $section1->addText('5.6. Hợp đồng có hiệu lực kể từ ngày ký và tự động hết hiệu lực khi hai bên hoàn thành nghĩa vụ của mình. Trong thời hạn 05 ngày kể từ ngày bên A hoàn thành nghĩa vụ thanh toán cho bên B, hợp đồng coi như được thanh lý.');
        $section1->addTextBreak();
        
        // Signatures
        $section1->addText('ĐẠI DIỆN BÊN A	                      	              ĐẠI DIỆN BÊN B', ['alignment' => 'center']);
        $section1->addTextBreak(8);
        
        // Tạo biên bản giao nhận hàng hóa (trang 3-4)
        $section2 = $phpWord->addSection();
        
        // Header
        $section2->addText('CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM', ['bold' => true, 'size' => 16], ['alignment' => 'center']);
        $section2->addText('Độc lập - Tự do - Hạnh phúc', ['bold' => true, 'size' => 14], ['alignment' => 'center']);
        $section2->addText('~~~~~o0o~~~~~', ['alignment' => 'center']);
        $section2->addText('BIÊN BẢN GIAO NHẬN HÀNG HOÁ KIÊM PHIẾU XUẤT KHO', ['bold' => true, 'size' => 14], ['alignment' => 'center']);
        $section2->addTextBreak();
        
        $section2->addText('Hôm nay, ngày    tháng   năm ' . $data['NamHienTai'] . ' tại Công ty TNHH sản xuất và thương mại dịch vụ Phú Thư , chúng tôi gồm có:');
        $section2->addTextBreak();
        
        // Section 1: Seller
        $section2->addText('1. BÊN A (BÊN BÁN) : CÔNG TY TNHH SẢN XUẤT VÀ THƯƠNG MẠI DỊCH VỤ PHÚ THƯ', ['bold' => true]);
        $section2->addText('Địa chỉ: Nhà số 1 thôn Rùa Thượng, Xã Thanh Thùy, Huyện Thanh Oai, Thành phố Hà Nội');
        $section2->addText('Điện thoại: 034.433.2918	Fax:');
        $section2->addText('Mã số thuế: 0110430598');
        $section2->addText('Đại diện: Bà Lý Ngọc Mai	Chức vụ: Kế toán');
        $section2->addTextBreak();
        
        // Section 2: Buyer
        $section2->addText('2. BÊN B (BÊN MUA) : ' . $data['Info_TenDonViDayDu'], ['bold' => true]);
        $section2->addText('Địa chỉ: ' . $data['Info_DiaChi']);
        $section2->addText('Điện thoại: ' . $data['Info_DienThoai'] . '	Fax: ' . $data['Info_Fax']);
        $section2->addText('Mã số thuế: ' . $data['Info_MaSoThue']);
        $section2->addText('Đại diện: …………………….	Chức vụ: ………………..');
        $section2->addTextBreak();
        
        $section2->addText('Hai bên thống nhất lập biên bản giao nhận hàng hoá với các nội dung sau:');
        $section2->addTextBreak();
        
        // Delivery Table
        $table2 = $section2->addTable(['borderSize' => 1]);
        
        // Header row
        $table2->addRow();
        $table2->addCell(1000)->addText('TT', ['bold' => true], ['alignment' => 'center']);
        $table2->addCell(3000)->addText('Tên hàng hóa', ['bold' => true], ['alignment' => 'center']);
        $table2->addCell(1500)->addText('Đơn vị tính', ['bold' => true], ['alignment' => 'center']);
        $table2->addCell(1500)->addText('Số lượng', ['bold' => true], ['alignment' => 'center']);
        $table2->addCell(2000)->addText('Đơn giá', ['bold' => true], ['alignment' => 'center']);
        $table2->addCell(2000)->addText('Thành tiền', ['bold' => true], ['alignment' => 'center']);
        
        // Empty rows
        $table2->addRow();
        $table2->addCell(1000)->addText('');
        $table2->addCell(3000)->addText('');
        $table2->addCell(1500)->addText('');
        $table2->addCell(1500)->addText('');
        $table2->addCell(2000)->addText('');
        $table2->addCell(2000)->addText('');
        
        $table2->addRow();
        $table2->addCell(1000)->addText('');
        $table2->addCell(3000)->addText('');
        $table2->addCell(1500)->addText('');
        $table2->addCell(1500)->addText('');
        $table2->addCell(2000)->addText('');
        $table2->addCell(2000)->addText('');
        
        // Data row
        $table2->addRow();
        $table2->addCell(1000)->addText('1', ['alignment' => 'center']);
        $table2->addCell(3000)->addText($data['TenHangHoaDichVu']);
        $table2->addCell(1500)->addText('Kg', ['alignment' => 'center']);
        $table2->addCell(1500)->addText($data['SoLuong'], ['alignment' => 'right']);
        $table2->addCell(2000)->addText($data['DonGia'], ['alignment' => 'right']);
        $table2->addCell(2000)->addText($data['ThanhTien'], ['alignment' => 'right']);
        
        // Summary rows
        $table2->addRow();
        $table2->addCell(9000)->addText('Giá trước thuế', ['bold' => true]);
        $table2->addCell(2000)->addText($data['ThanhTien'], ['bold' => true, 'alignment' => 'right']);
        
        $table2->addRow();
        $table2->addCell(9000)->addText('Thuế GTGT 10%', ['bold' => true]);
        $table2->addCell(2000)->addText($data['TienThueGTGT'], ['bold' => true, 'alignment' => 'right']);
        
        $table2->addRow();
        $table2->addCell(9000)->addText('Tổng cộng', ['bold' => true]);
        $table2->addCell(2000)->addText($data['TongTienThanhToan'], ['bold' => true, 'alignment' => 'right']);
        
        $section2->addTextBreak();
        $section2->addText('Số tiền bằng chữ: ' . $data['TongTienThanhToan_BangChu']);
        $section2->addTextBreak();
        $section2->addText('Biên bản này được lập thành 02 bản có giá trị pháp lý như nhau, bên A giữ 01 bản, bên B giữ 01 bản .');
        $section2->addTextBreak();
        
        // Signatures
        $section2->addText('ĐẠI DIỆN BÊN BÁN                                                            ĐẠI DIỆN BÊN MUA', ['alignment' => 'center']);
        $section2->addTextBreak(8);
        
        // Lưu file Word
        $writer = WordIOFactory::createWriter($phpWord, 'Word2007');
        $writer->save($outputPath);
        
        return true;
    } catch (Exception $e) {
        throw $e;
    }
}

function createContractContent($phpWord, $data) {
    $section = $phpWord->getSections()[0];
        
        // Header
        $section->addText('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', ['bold' => true, 'size' => 16], ['alignment' => 'center']);
        $section->addText('Độc lập - Tự do - Hạnh phúc', ['bold' => true, 'size' => 14], ['alignment' => 'center']);
        $section->addText('---- o0o ----', ['alignment' => 'center']);
        $section->addText('HỢP ĐỒNG MUA BÁN', ['bold' => true, 'size' => 14], ['alignment' => 'center']);
        $section->addTextBreak();
        
        // Contract Number
        $section->addText('Số: ' . $data['NgayThangNamHD'] . '/HĐMB/PT-' . $data['Info_TenCongTy'], ['bold' => true]);
        $section->addTextBreak();
        
        // Legal Basis
        $section->addText('- Căn cứ vào các Bộ luật dân sự Nước Cộng hòa Xã hội Chủ nghĩa Việt Nam');
        $section->addText('- Căn cứ vào Luật Thương mại 2005 của Nước Cộng hòa Xã hội Chủ nghĩa Việt Nam');
        $section->addText('- Căn cứ vào nhu cầu và khả năng cung cấp hàng hóa của các bên.');
        $section->addTextBreak();
        
        // Date and Parties Introduction
        $section->addText('Hôm nay, ngày    tháng    năm ' . $data['NamHienTai'] . ' chúng tôi gồm:');
        $section->addTextBreak();
        
        // Section I: Seller
        $section->addText('I. BÊN BÁN (BÊN A): CÔNG TY TNHH SẢN XUẤT VÀ THƯƠNG MẠI DỊCH VỤ PHÚ THƯ', ['bold' => true]);
        $section->addText('Địa chỉ: Nhà số 1 Thôn Rùa Thượng – Xã Thanh Thùy – Huyện Thanh Oai - Thành phố Hà Nội, Việt Nam');
        $section->addText('Điện thoại: 034.433.2918	Fax:');
        $section->addText('Mã số thuế: 0110430598');
        $section->addText('Đại diện: Ông Nguyễn Văn Báo           Chức vụ: Giám đốc');
        $section->addText('Tài khoản: 117002948428 – Tại Ngân hàng Vietinbank – Chi nhánh Thành An');
        $section->addTextBreak();
        
        // Section II: Buyer
        $section->addText('II. BÊN MUA ( BÊN B): ' . $data['Info_TenDonViDayDu'], ['bold' => true]);
        $section->addText('Địa chỉ: ' . $data['Info_DiaChi']);
        $section->addText('Điện thoại: ' . $data['Info_DienThoai'] . '	Fax: ' . $data['Info_Fax']);
        $section->addText('Mã số thuế: ' . $data['Info_MaSoThue']);
        $section->addText('Đại diện: ' . $data['Info_DaiDien'] . '	Chức vụ: ' . $data['Info_ChucVu']);
        $section->addText('Tài khoản: ' . $data['Info_TaiKhoan']);
        $section->addTextBreak();
        
        // Agreement Statement
        $section->addText('Sau khi thỏa thuận, bên B đồng ý mua, Bên A đồng ý bán hàng hóa theo các điều kiện sau đây:');
        $section->addTextBreak();
        
        // Article 1: Goods, Quality and Quantity
        $section->addText('Điều 1: Hàng hóa, chất lượng và số lượng.', ['bold' => true]);
        $section->addText('1.1. Hàng hóa và số lượng: Bên B đồng ý mua của bên A mặt hàng ' . $data['TenHangHoaDichVu'] . ' do bên A bán với số lượng chi tiết ở điều 2.1 của hợp đồng.');
        $section->addText('1.2. Chất lượng: Hàng hóa mới 100%.');
        $section->addTextBreak(3);
        
        // Article 2: Price and Payment Method
        $section->addText('Điều 2: Giá cả, phương thức thanh toán.', ['bold' => true]);
        $section->addText('2.1. Giá cả: giá trị của hàng hóa nêu tại Điều 1 của hợp đồng này được tính theo bảng giá như sau:');
        
        // Price Table
        $table = $section->addTable(['borderSize' => 1]);
        
        // Header row
        $table->addRow();
        $table->addCell(1000)->addText('STT', ['bold' => true], ['alignment' => 'center']);
        $table->addCell(3000)->addText('LOẠI HÀNG', ['bold' => true], ['alignment' => 'center']);
        $table->addCell(1000)->addText('ĐVT', ['bold' => true], ['alignment' => 'center']);
        $table->addCell(1500)->addText('S.LƯỢNG', ['bold' => true], ['alignment' => 'center']);
        $table->addCell(2000)->addText('ĐƠN GIÁ', ['bold' => true], ['alignment' => 'center']);
        $table->addCell(2000)->addText('THÀNH TIỀN', ['bold' => true], ['alignment' => 'center']);
        
        // Data row
        $table->addRow();
        $table->addCell(1000)->addText('1', ['alignment' => 'center']);
        $table->addCell(3000)->addText($data['TenHangHoaDichVu']);
        $table->addCell(1000)->addText('Kg', ['alignment' => 'center']);
        $table->addCell(1500)->addText($data['SoLuong'], ['alignment' => 'right']);
        $table->addCell(2000)->addText($data['DonGia'], ['alignment' => 'right']);
        $table->addCell(2000)->addText($data['ThanhTien'], ['alignment' => 'right']);
        
        // Summary rows
        $table->addRow();
        $table->addCell(8500)->addText('Tổng cộng tiền hàng', ['bold' => true]);
        $table->addCell(2000)->addText($data['ThanhTien'], ['bold' => true, 'alignment' => 'right']);
        
        $table->addRow();
        $table->addCell(8500)->addText('Thuế VAT 10%', ['bold' => true]);
        $table->addCell(2000)->addText($data['TienThueGTGT'], ['bold' => true, 'alignment' => 'right']);
        
        $table->addRow();
        $table->addCell(8500)->addText('Tổng cộng', ['bold' => true]);
        $table->addCell(2000)->addText($data['TongTienThanhToan'], ['bold' => true, 'alignment' => 'right']);
        
        $section->addTextBreak();
        $section->addText('Bằng chữ: ' . $data['TongTienThanhToan_BangChu']);
        $section->addText('Giá trị hợp đồng được phép dung sai tối đa +/-10%');
        $section->addTextBreak();
        
        $section->addText('2.2. Phương thức thanh toán:', ['bold' => true]);
        $section->addText('Hình thức thanh toán: Thanh toán trước hoặc sau không quá 60 ngày khi bên A giao đủ hàng và chứng từ cho bên B, bằng tiền mặt hoặc chuyển khoản.');
        $section->addText('Chứng từ thanh toán bao gồm: Hoá đơn GTGT, Biên bản giao nhận hàng hóa.');
        $section->addTextBreak();
        
        // Article 3: Delivery Location and Method
        $section->addText('Điều 3: Địa điểm và phương thức giao nhận hàng hóa', ['bold' => true]);
        $section->addText('3.1. Địa điểm giao nhận:', ['bold' => true]);
        $section->addText('Tại xưởng bán bên A');
        $section->addText('3.2. Phương thức giao nhận:', ['bold' => true]);
        $section->addText('Giao hàng qua cân thực tế và lập Biên bản giao nhận hàng hóa giữa hai bên. Khối lượng trên biên bản được xác định làm căn cứ thanh toán. Nếu phát hiện hàng thiếu hoặc không đúng chất lượng, bên B có quyền từ chối nhận hàng và yêu cầu bên A giao lại hàng đúng theo hợp đồng.');
        $section->addTextBreak();
        
        // Article 4: Responsibilities
        $section->addText('Điều 4: Trách nhiệm các bên', ['bold' => true]);
        $section->addText('4.1. Trách nhiệm bên A:', ['bold' => true]);
        $section->addText('- Giao hàng đúng số lượng, chất lượng, thời gian và địa điểm đã thỏa thuận.');
        $section->addText('- Cung cấp đầy đủ chứng từ pháp lý về hàng hóa.');
        $section->addText('- Chịu trách nhiệm về chất lượng hàng hóa trong thời gian bảo hành.');
        $section->addText('- Đảm bảo hàng hóa đúng tiêu chuẩn chất lượng đã cam kết.');
        $section->addText('- Hỗ trợ bên B trong việc vận chuyển và bốc dỡ hàng hóa.');
        $section->addText('4.2. Trách nhiệm bên B:', ['bold' => true]);
        $section->addText('- Thanh toán đúng thời hạn và phương thức đã thỏa thuận.');
        $section->addText('- Nhận hàng đúng thời gian và địa điểm đã thỏa thuận.');
        $section->addText('- Kiểm tra hàng hóa trước khi nhận.');
        $section->addText('- Thông báo kịp thời cho bên A nếu có vấn đề về chất lượng hàng hóa.');
        $section->addText('- Chịu trách nhiệm về việc bảo quản hàng hóa sau khi nhận.');
        $section->addTextBreak();
        
        // Article 5: Delivery Time
        $section->addText('Điều 5: Thời gian giao hàng', ['bold' => true]);
        $section->addText('5.1. Thời gian giao hàng:', ['bold' => true]);
        $section->addText('Bên A sẽ giao hàng cho bên B trong vòng 15 ngày kể từ ngày ký hợp đồng hoặc theo thỏa thuận cụ thể giữa hai bên.');
        $section->addText('5.2. Gia hạn giao hàng:', ['bold' => true]);
        $section->addText('Trong trường hợp bất khả kháng, bên A có thể gia hạn thời gian giao hàng nhưng phải thông báo trước cho bên B ít nhất 3 ngày.');
        $section->addTextBreak();
        
        // Article 6: Warranty
        $section->addText('Điều 6: Bảo hành', ['bold' => true]);
        $section->addText('6.1. Thời gian bảo hành:', ['bold' => true]);
        $section->addText('Bên A cam kết bảo hành hàng hóa trong vòng 12 tháng kể từ ngày giao hàng.');
        $section->addText('6.2. Điều kiện bảo hành:', ['bold' => true]);
        $section->addText('- Hàng hóa phải được sử dụng đúng mục đích và hướng dẫn.');
        $section->addText('- Không được sửa chữa, thay đổi cấu trúc hàng hóa.');
        $section->addText('- Có hóa đơn và phiếu bảo hành hợp lệ.');
        $section->addTextBreak();
        
        // Article 7: Force Majeure
        $section->addText('Điều 7: Bất khả kháng', ['bold' => true]);
        $section->addText('7.1. Định nghĩa:', ['bold' => true]);
        $section->addText('Bất khả kháng là các sự kiện nằm ngoài khả năng kiểm soát của các bên như thiên tai, hỏa hoạn, chiến tranh, đình công, các quy định của cơ quan nhà nước.');
        $section->addText('7.2. Xử lý:', ['bold' => true]);
        $section->addText('Khi xảy ra bất khả kháng, các bên sẽ thương lượng để tìm giải pháp phù hợp và không bên nào chịu trách nhiệm bồi thường.');
        $section->addTextBreak();
        
        // Article 8: Dispute Resolution
        $section->addText('Điều 8: Giải quyết tranh chấp', ['bold' => true]);
        $section->addText('8.1. Thương lượng:', ['bold' => true]);
        $section->addText('Mọi tranh chấp phát sinh từ hợp đồng này sẽ được giải quyết thông qua thương lượng hòa bình giữa các bên.');
        $section->addText('8.2. Trọng tài:', ['bold' => true]);
        $section->addText('Trong trường hợp không thương lượng được, tranh chấp sẽ được đưa ra Tòa án có thẩm quyền tại Hà Nội để giải quyết.');
        $section->addTextBreak();
        
        // Article 9: Contract Validity
        $section->addText('Điều 9: Hiệu lực hợp đồng', ['bold' => true]);
        $section->addText('Hợp đồng này có hiệu lực từ ngày ký và được lập thành 02 (hai) bản có giá trị pháp lý như nhau, mỗi bên giữ 01 (một) bản.');
        $section->addTextBreak();
        
        // Article 10: Amendments
        $section->addText('Điều 10: Sửa đổi, bổ sung', ['bold' => true]);
        $section->addText('Mọi sửa đổi, bổ sung hợp đồng này phải được thực hiện bằng văn bản và có chữ ký của đại diện có thẩm quyền của cả hai bên.');
        $section->addTextBreak(2);
        
        // Signatures
        $section->addText('BÊN BÁN', ['bold' => true], ['alignment' => 'left']);
        $section->addText('(Ký tên, đóng dấu)', ['alignment' => 'left']);
        $section->addTextBreak(3);
        $section->addText('BÊN MUA', ['bold' => true], ['alignment' => 'right']);
        $section->addText('(Ký tên, đóng dấu)', ['alignment' => 'right']);
        
        // Lưu file Word
        $writer = WordIOFactory::createWriter($phpWord, 'Word2007');
        $writer->save($outputPath);
        
        return true;
    } catch (Exception $e) {
        throw $e;
    }
}

// 📌 Xử lý HĐMB với HTML template
if ($docType === 'HĐMB_HTML') {
    foreach ($data as $row) {
        // 👉 Xử lý công ty trước để chọn đúng template
        $maSoThue = $row['L'] ?? '';
        if (!isset($companies[$maSoThue])) {
            echo "<h2 style='color: red;'>❌ Không tìm thấy công ty có mã số thuế: <strong>$maSoThue</strong></h2>";
            echo "<a href='index.php'><button>⬅ Quay lại</button></a>";
            exit;
        }

        $companyInfo = $companies[$maSoThue];

        //-------------------------------------------------------------------------------
        // 👉 Xử lý đặc biệt cột E - NgayThangNamHD
        $originalDate = $row['E'] ?? '';
        $formattedDate = '';
        $originalYear = '';

        if (!empty($originalDate)) {
            if (is_numeric($originalDate)) {
                $phpDate = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($originalDate);
            } else {
                $phpDate = DateTime::createFromFormat('d/m/Y', $originalDate) ?: DateTime::createFromFormat('d-m-Y', $originalDate);
            }

            if ($phpDate) {
                $phpDate->modify('-10 days');
                $formattedDate = $phpDate->format('dmY');
                $originalYear = $phpDate->format('Y');
            }
        }

        $row['E'] = $formattedDate;

        // 👉 Chuẩn bị dữ liệu cho Word document
        $fieldsNeedFormat = ['SoLuong', 'DonGia', 'ThanhTien', 'TienThueGTGT', 'TongTienThanhToan'];
        $tongTienNumeric = 0;

        // Xử lý tổng tiền thanh toán
        $tongTienCol = array_search('TongTienThanhToan', $headers);
        if ($tongTienCol && isset($row[$tongTienCol])) {
            $valueClean = str_replace(',', '', $row[$tongTienCol]);
            $tongTienNumeric = (float)$valueClean;
        }

        // Xử lý dữ liệu an toàn - làm sạch ký tự đặc biệt
        $safeValue = function($value) {
            if ($value === null || $value === '') {
                return '';
            }
            // Loại bỏ ký tự không in được và escape HTML
            $cleanValue = preg_replace('/[\x00-\x1F\x7F-\x9F]/', '', (string)$value);
            return htmlspecialchars($cleanValue, ENT_QUOTES, 'UTF-8');
        };
        
        // Map dữ liệu theo đúng template DOCX gốc - đầy đủ các field
        $tongTienNumeric = (float)str_replace(',', '', $safeValue($row['S']));
        $documentData = [
            'Info_TenDonViDayDu' => $safeValue($companyInfo['ten']),
            'Info_MaSoThue' => $safeValue($maSoThue),
            'Info_DiaChi' => $safeValue($companyInfo['diachi']),
            'Info_DienThoai' => $safeValue($companyInfo['sdt']),
            'Info_Fax' => $safeValue($companyInfo['fax']),
            'Info_DaiDien' => $safeValue($companyInfo['daidien']),
            'Info_ChucVu' => $safeValue($companyInfo['chucvu']),
            'Info_TaiKhoan' => $safeValue($companyInfo['taikhoan']),
            'Info_TenCongTy' => $safeValue($companyInfo['ten']),
            'NgayThangNamHD' => $formattedDate,
            'TenHangHoaDichVu' => $safeValue($row['M']),
            'SoLuong' => number_format((float)str_replace(',', '', $safeValue($row['N'])), 0, ',', '.'),
            'DonGia' => number_format((float)str_replace(',', '', $safeValue($row['O'])), 0, ',', '.'),
            'ThanhTien' => number_format((float)str_replace(',', '', $safeValue($row['P'])), 0, ',', '.'),
            'TienThueGTGT' => number_format((float)str_replace(',', '', $safeValue($row['R'])), 0, ',', '.'),
            'TongTienThanhToan' => number_format($tongTienNumeric, 0, ',', '.'),
            'TongTienThanhToan_BangChu' => numberToVietnameseWords($tongTienNumeric),
            'NamHienTai' => $originalYear ?? ''
        ];

        //👉 Lưu file vào thư mục
        $codeName = $companyInfo['code_name'];
        $companyDir = $outputDir . $codeName . '/';
        if (!is_dir($companyDir)) {
            mkdir($companyDir, 0777, true);
        }

        $fileName = "HĐMB_{$originalYear}_{$codeName}_{$index}.docx";
        $savePath = $companyDir . $fileName;
        
        try {
            createWordDirectly($documentData, $savePath);
            
            // Kiểm tra file đã được tạo
            if (file_exists($savePath) && filesize($savePath) > 0) {
                echo "<div style='color: green;'>✅ Tạo thành công: $fileName (" . filesize($savePath) . " bytes)</div>";
            } else {
                echo "<div style='color: red;'>❌ File không được tạo hoặc rỗng: $fileName</div>";
            }
        } catch (Exception $e) {
            echo "<div style='color: red;'>❌ Lỗi khi tạo file Word: " . $e->getMessage() . "</div>";
            echo "<div style='color: red;'>Stack trace: " . $e->getTraceAsString() . "</div>";
            continue;
        }

        $index++;
    }

    echo "<div style='
        font-family: Arial, sans-serif;
        background-color: #d4edda;
        color: #155724;
        padding: 20px;
        margin: 50px auto;
        border: 1px solid #c3e6cb;
        border-radius: 10px;
        width: fit-content;
        text-align: center;
    '>
        <p><strong>✅ Thành công!</strong> Đã tạo " . ($index - 1) . " hợp đồng từ HTML template tại thư mục <code>output</code>.</p>
        <div style='margin-top: 15px;'>
            <a href='output/' target='_blank' style='
                display: inline-block;
                padding: 10px 20px;
                margin-right: 10px;
                background-color: #28a745;
                color: white;
                text-decoration: none;
                border-radius: 6px;
                font-weight: bold;
            '>📂 Xem hợp đồng</a>

            <a href='index.php' style='
                display: inline-block;
                padding: 10px 20px;
                background-color: #007bff;
                color: white;
                text-decoration: none;
                border-radius: 6px;
                font-weight: bold;
            '>⬅️ Quay về</a>
        </div>
    </div>";
    exit;
}

// 📌 Xử lý HĐMB (Xuất file word mỗi dòng)
foreach ($data as $row) {
    // 👉 Xử lý công ty trước để chọn đúng template
    $maSoThue = $row['L'] ?? '';
    if (!isset($companies[$maSoThue])) {
        echo "<h2 style='color: red;'>❌ Không tìm thấy công ty có mã số thuế: <strong>$maSoThue</strong></h2>";
        echo "<a href='index.php'><button>⬅ Quay lại</button></a>";
        exit;
    }

    $companyInfo = $companies[$maSoThue];
    $isThaiDuong = $companyInfo['code'] === 'TD';

    // 👉 Chọn template
    $templateFile = $isThaiDuong ? 'template2.docx' : 'template.docx';
    $template = new TemplateProcessor($templateFile);

    //-------------------------------------------------------------------------------
    // 👉 Xử lý đặc biệt cột E - NgayThangNamHD
    $originalDate = $row['E'] ?? '';
    $formattedDate = '';
    $originalYear = '';

    if (!empty($originalDate)) {
        if (is_numeric($originalDate)) {
            $phpDate = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($originalDate);
        } else {
            $phpDate = DateTime::createFromFormat('d/m/Y', $originalDate) ?: DateTime::createFromFormat('d-m-Y', $originalDate);
        }

        if ($phpDate) {
            $phpDate->modify('-10 days');
            $formattedDate = $phpDate->format('dmY');
            $originalYear = $phpDate->format('Y');
        }
    }

    $row['E'] = $formattedDate;

    //-------------------------------------------------------------------------------
    // 👉 Gán thông tin công ty
    $template->setValue('Info_MaSoThue',         $maSoThue);
    $template->setValue('Info_TenCongTy',        $companyInfo['code']);
    $template->setValue('Info_TenDonViDayDu',    $companyInfo['ten']);
    $template->setValue('Info_DiaChi',           $companyInfo['diachi']);
    $template->setValue('Info_DienThoai',        $companyInfo['sdt']);
    $template->setValue('Info_Fax',              $companyInfo['fax']);
    $template->setValue('Info_DaiDien',          $companyInfo['daidien']);
    $template->setValue('Info_ChucVu',           $companyInfo['chucvu']);

    // 👉 Tài khoản: xử lý riêng cho TD
    if ($isThaiDuong) {
        $taikhoanLines = preg_split('/\r\n|\r|\n/', trim($companyInfo['taikhoan']));
        $template->setValue('Info_TaiKhoan1', $taikhoanLines[0] ?? '');
        $template->setValue('Info_TaiKhoan2', $taikhoanLines[1] ?? '');
    } else {
        $template->setValue('Info_TaiKhoan', trim($companyInfo['taikhoan']));
    }

    // 👉 Gán năm vào template
    $template->setValue('NamHienTai', $originalYear ?? '');

    // 👉 Map dữ liệu từ Excel
    $fieldsNeedFormat = ['SoLuong', 'DonGia', 'ThanhTien', 'TienThueGTGT', 'TongTienThanhToan'];
    $tongTienFormatted = '';
    $tongTienNumeric = 0;

    foreach ($headers as $col => $fieldName) {
        $value = $row[$col] ?? '';

        if (in_array($fieldName, $fieldsNeedFormat)) {
            $valueClean = str_replace(',', '', $value);
            $formatted = number_format((float)$valueClean, 0, ',', '.');

            if ($fieldName === 'TongTienThanhToan') {
                $tongTienFormatted = $formatted;
                $tongTienNumeric = (float)$valueClean;
            }

            $value = $formatted;
        }

        $template->setValue($fieldName, $value);
    }

    // 👉 Thêm giá trị bằng chữ cho TongTienThanhToan
    $tienBangChu = numberToVietnameseWords($tongTienNumeric);
    $template->setValue('TongTienThanhToan_BangChu', $tienBangChu);

    //👉 Lưu file vào thư mục
    $codeName = $companyInfo['code_name'];
    $companyDir = $outputDir . $codeName . '/';
    if (!is_dir($companyDir)) {
        mkdir($companyDir, 0777, true);
    }

    $fileName = "HĐMB_{$originalYear}_{$codeName}_{$index}.docx";
    $savePath = $companyDir . $fileName;
    $template->saveAs($savePath);

    $index++;
}


echo "<div style='
    font-family: Arial, sans-serif;
    background-color: #d4edda;
    color: #155724;
    padding: 20px;
    margin: 50px auto;
    border: 1px solid #c3e6cb;
    border-radius: 10px;
    width: fit-content;
    text-align: center;
'>
    <p><strong>✅ Thành công!</strong> Đã tạo " . ($index - 1) . " hợp đồng tại thư mục <code>output</code>.</p>
    <div style='margin-top: 15px;'>
        <a href='output/' target='_blank' style='
            display: inline-block;
            padding: 10px 20px;
            margin-right: 10px;
            background-color: #28a745;
            color: white;
            text-decoration: none;
            border-radius: 6px;
            font-weight: bold;
        '>📂 Xem hợp đồng</a>

        <a href='index.php' style='
            display: inline-block;
            padding: 10px 20px;
            background-color: #007bff;
            color: white;
            text-decoration: none;
            border-radius: 6px;
            font-weight: bold;
        '>⬅️ Quay về</a>
    </div>
</div>";
