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

/**
 * Tên viết tắt của công ty, rút về dạng đặt được cho thư mục trên Windows.
 */
function folderName(array $companyInfo) {
    $name = $companyInfo['ten_viet_tat'] ?? str_replace('_', ' ', $companyInfo['code_name']);
    $name = preg_replace('/[\\\\\/:*?"<>|]+/u', ' ', $name);
    $name = preg_replace('/\s+/u', ' ', trim($name));
    return mb_substr($name, 0, 90);
}

/**
 * Thư mục riêng của một công ty trong lần export này: {TenCongTy}_{ngay-gio-phut}.
 * Export lại trong cùng một phút thì tách ra -2, -3... để không đè file lần trước.
 */
function makeExportDir($baseDir, $codeName, $stamp) {
    $name = $codeName . '_' . $stamp;

    $path = $baseDir . $name;
    for ($n = 2; is_dir($path); $n++) {
        $path = $baseDir . $name . '-' . $n;
    }

    mkdir($path, 0777, true);
    return $path . '/';
}

// 👉 Mỗi loại tài liệu một thư mục, bên trong mỗi công ty một thư mục kèm thời điểm export
$docFolders = [
    'HĐMB'      => 'HopDongMuaBan',
    'HĐMB_HTML' => 'HopDongMuaBan',
    'BBGN'      => 'BienBanGiaoNhan_Excel',
    'BBGN_WORD' => 'BienBanGiaoNhan_Word',
    'DDH'       => 'DonDatHang',
];

$exportFolder = $docFolders[$docType] ?? 'Khac';
$exportBase   = $outputDir . $exportFolder . '/';
$exportLink   = 'output/' . $exportFolder . '/';
$exportStamp  = date('d-m-Y_H-i');
$companyDirs  = [];

/**
 * Link hiển thị sau khi export xong: chỉ có một công ty thì trỏ thẳng vào thư mục
 * của công ty đó, nhiều công ty thì trỏ vào thư mục của loại tài liệu.
 */
function exportLink($exportLink, array $companyDirs) {
    if (count($companyDirs) !== 1) return $exportLink;
    return $exportLink . basename(reset($companyDirs)) . '/';
}

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
        'ten'      => 'CÔNG TY CỔ PHẦN SẢN XUẤT THIẾT BỊ ĐIỆN KING',
        'code'     => 'K',
        'code_name'=> 'KING',
        'ten_viet_tat' => 'King',
        'diachi'   => 'Thôn Từ Am, Xã Tam Hưng, Thành phố Hà Nội, Việt Nam',
        'sdt'      => '097.499.1127',
        'fax'      => '',
        'daidien'  => 'Ông Cao Anh Tuấn',
        'chucvu'   => 'Giám đốc',
        'taikhoan' => '1121166699999 tại Ngân hàng TMCP Quân Đội - Chi nhánh Sở giao dịch 1',
    ],
    '0104315188' => [
        'ten'      => 'CÔNG TY CỔ PHẦN SẢN XUẤT CƠ KHÍ VÀ THƯƠNG MẠI THÁI DƯƠNG',
        'code'     => 'TD',
        'code_name'=> 'THAI_DUONG',
        'ten_viet_tat' => 'Thái Dương',
        'diachi'   => 'Thôn Từ Am, Xã Tam Hưng, Thành phố Hà Nội, Việt Nam',
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
        'ten_viet_tat' => 'Tuấn Kiệt',
        'diachi'   => 'Số nhà 35, ngõ Đồng Đanh, thôn Rùa Thượng, xã Tam Hưng,Thành Phố Hà Nội, Việt Nam.',
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
        'ten_viet_tat' => 'Tiến Thành',
        'diachi'   => 'Thôn Rùa Hạ, Xã Tam Hưng, Thành phố Hà Nội, Việt Nam',
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
        'ten_viet_tat' => 'Hoàng Kim',
        'diachi'   => 'Số 6, Xóm Mới, Thôn Rùa Thượng, Xã Tam Hưng, Thành phố Hà Nội, Việt Nam',
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
        'ten_viet_tat' => 'Minh Tâm',
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
        'ten_viet_tat' => 'Nhật Minh',
        'diachi'   => 'Thôn Rùa Thượng, Xã Tam Hưng, Thành phố Hà Nội, Việt Nam',
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
        'ten_viet_tat' => 'Hanoime',
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
        'ten_viet_tat' => 'Đại Linh',
        'diachi'   => 'Tổ 15, Thị trấn Nam Giang, Huyện Nam Trực, Tỉnh Nam Định, Việt Nam',
        'sdt'      => '',
        'fax'      => '',
        'daidien'  => '',
        'chucvu'   => '',
        'taikhoan' => '',
    ],
    '0109691813' => [
        'ten'      => 'CÔNG TY CỔ PHẦN SẢN XUẤT VÀ THƯƠNG MẠI ANH QUÂN',
        'code'     => 'AQ',
        'code_name'=> 'ANH_QUAN',
        'ten_viet_tat' => 'Anh Quân',
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
        'ten_viet_tat' => 'Kiên Gia',
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
        'ten_viet_tat' => 'Thiên Phú',
        'diachi'   => 'Thôn Rùa Hạ, xã Tam Hưng, thành phố Hà Nội',
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
       'ten_viet_tat' => 'Khang Thịnh Phát',
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
        'ten_viet_tat' => 'DMEC',
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
        'ten_viet_tat' => 'PCCC Gia Phát',
        'diachi'   => 'Thôn Hoa Thám, Xã An Khánh, Thành phố Hà Nội, Việt Nam',
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
        'ten_viet_tat' => 'Hoàng Minh Mechanical',
        'diachi'   => 'Thôn Rùa Hạ, Xã Tam Hưng, Thành phố Hà Nội, Việt Nam',
        'sdt'      => '',
        'fax'      => '',
        'daidien'  => 'Ông Thái Đình Kiên',
        'chucvu'   => 'Giám đốc',
        'taikhoan' => '',
    ],
	'0106029785' => [
        'ten'      => 'CÔNG TY CỔ PHẦN THƯƠNG MẠI VÀ THIẾT BỊ TRƯỜNG AN',
        'code'     => 'TA',
        'code_name'=> 'Truong_An',
        'ten_viet_tat' => 'Trường An',
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
        'ten_viet_tat' => 'Anh Linh',
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
        'ten_viet_tat' => 'Phong Thái',
        'diachi'   => 'Thôn Rùa Hạ, Xã Tam Hưng, Thành phố Hà Nội, Việt Nam',
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
        'ten_viet_tat' => 'Hợp Phát',
        'diachi'   => 'Số nhà 4 nghách 123 ngõ 324, đường Phương Canh, Phường Xuân Phương, TP Hà Nội, Việt Nam',
        'sdt'      => '',
        'fax'      => '',
        'daidien'  => 'TRẦN TRUNG TUYẾN',
        'chucvu'   => '',
        'taikhoan' => '4520625249 tại BIDV CN Vạn Phúc - Hà Nội',
    ],
    '0107017609' => [
        'ten'      => 'CÔNG TY TNHH SẢN XUẤT VÀ PHÁT TRIỂN THƯƠNG MẠI MINH TÂM',
        'code'     => 'MT',
        'code_name'=> 'Minh_Tam_TuAm',
        'ten_viet_tat' => 'Minh Tâm Từ Am',
        'diachi'   => 'Thôn Từ Am, Xã Tam Hưng, Tp Hà Nội, Việt Nam',
        'sdt'      => '098.126.8969',
        'fax'      => '',
        'daidien'  => 'Ông Nguyễn Văn Tâm',
        'chucvu'   => 'Giám Đốc',
        'taikhoan' => '020038194888 tại ngân hàng Sacombank - PGD Thường Tín - CN Thường Tín',
    ],
    '0500571361' => [
        'ten'      => 'CÔNG TY TNHH SẢN XUẤT VÀ THƯƠNG MẠI KIM KHÍ TÂN PHÁT',
        'code'     => 'TP',
        'code_name'=> 'Tan_Phat',
        'ten_viet_tat' => 'Tân Phát',
        'diachi'   => 'Thôn Rùa Hạ, Xã Tam Hưng, Thành phố Hà Nội, Việt Nam',
        'sdt'      => '',
        'fax'      => '',
        'daidien'  => 'Bà Đinh Thị Minh Thu',
        'chucvu'   => 'Giám Đốc',
        'taikhoan' => '24631656 tại Ngân Hàng Việt Nam Thịnh Vượng VPBank - Chi Nhánh Hà Đông',
    ],
    '0108957179' => [
        'ten'      => 'CÔNG TY TNHH TYAK Việt Nam',
        'code'     => 'TYAK',
        'code_name'=> 'TYAK',
        'ten_viet_tat' => 'TYAK',
        'diachi'   => 'Thôn Rùa Hạ, Xã Tam Hưng, Thành phố Hà Nội, Việt Nam',
        'sdt'      => '',
        'fax'      => '',
        'daidien'  => 'Ông Nguyễn Đức Toàn',
        'chucvu'   => 'Giám Đốc',
        'taikhoan' => '0728669766666 tại Ngân Hàng TMCP quân đội MB - Chi nhánh Thanh Trì',
    ],
    '0111518365' => [
        'ten'      => 'CÔNG TY TNHH SẢN XUẤT VÀ THƯƠNG MẠI ANH TRIẾT',
        'code'     => 'AT',
        'code_name'=> 'AT',
        'ten_viet_tat' => 'Anh Triết',
        'diachi'   => 'Thôn Rùa Hạ, Xã Tam Hưng, Thành phố Hà Nội, Việt Nam',
        'sdt'      => '098.973.3950',
        'fax'      => '',
        'daidien'  => 'Ông Hoàng Văn Tới',
        'chucvu'   => 'Giám Đốc',
        'taikhoan' => '90972325 tại Ngân Hàng TMCP quân đội MB - Chi nhánh Thụy Khuê',
    ],
    '0111167893' => [
        'ten'      => 'CÔNG TY TNHH PCCC TRẦN GIA',
        'code'     => 'TG',
        'code_name'=> 'PCCC_TRAN_GIA',
        // Tên viết tắt dùng trong số hợp đồng nguyên tắc; bỏ trống thì lấy code_name
        'ten_viet_tat' => 'PCCC Trần Gia',
        'diachi'   => 'Số nhà 8, Ngách 7/24, Ngõ 311, Xóm Tiền Phong, Xã An Khánh, TP Hà Nội, Việt Nam',
        'sdt'      => '',
        'fax'      => '',
        'daidien'  => 'Bà Trần Thị Thanh Huyền',
        'chucvu'   => 'Giám đốc',
        'taikhoan' => '2299686868 tại ngân hàng TMCP Sài Gòn - Hà Nội – Chi nhánh Hà Thành',
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

/**
 * Ngày trong Excel có thể là số serial hoặc chuỗi d/m/Y, d-m-Y
 * @return DateTime|null
 */
function parseExcelDate($value) {
    if ($value === null || $value === '') return null;

    if (is_numeric($value)) {
        return \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($value);
    }

    $value = trim((string)$value);
    $date = DateTime::createFromFormat('d/m/Y', $value) ?: DateTime::createFromFormat('d-m-Y', $value);
    return $date ?: null;
}

/** "  53,510.000000 " => 53510.0 */
function toNumber($value) {
    return (float) str_replace(',', '', trim((string)$value));
}

/** "  53,510.000000 " => "53.510" */
function formatNumber($value) {
    return number_format(toNumber($value), 0, ',', '.');
}

/**
 * Chuẩn bị giá trị trước khi đổ vào template Word:
 * escape ký tự XML (tên công ty có dấu &) và giữ lại chỗ xuống dòng.
 */
function wordValue($value) {
    $escaped = htmlspecialchars(trim((string)$value), ENT_NOQUOTES | ENT_XML1, 'UTF-8');
    return preg_replace('/\s*\R\s*/u', '</w:t><w:br/><w:t xml:space="preserve">', $escaped);
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

        if (!isset($companyDirs[$codeName])) {
            $companyDirs[$codeName] = makeExportDir($exportBase, folderName($companyInfo), $exportStamp);
        }
        $companyDir = $companyDirs[$codeName];

        $ymd = $phpDate->format('d_m_Y');
        $fileName = "BBGN_{$ymd}_{$codeName}_{$index}.xlsx";
        $savePath = $companyDir . $fileName;
        $writer = IOFactory::createWriter($templateExcel, 'Xlsx');
        $writer->save($savePath);

        $index++;
    }

    $ketQuaLink = exportLink($exportLink, $companyDirs);
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
        <p><strong>✅ Thành công!</strong> Đã tạo file BBGN trong thư mục <code>{$ketQuaLink}</code>.</p>
        <div style='margin-top: 15px;'>
            <a href='{$ketQuaLink}' target='_blank' style='
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

// 📌 Xử lý BBGN (Word) và ĐĐH — cùng nguồn dữ liệu, chỉ khác template và nơi lưu
if ($docType === 'BBGN_WORD' || $docType === 'DDH') {
    $isDonDatHang = ($docType === 'DDH');

    $config = $isDonDatHang
        ? ['template' => 'template_DDH.docx',       'prefix' => 'DDH']
        : ['template' => 'template_BBGN_word.docx', 'prefix' => 'BBGN'];

    if (!file_exists($config['template'])) {
        echo "<h2 style='color: red;'>❌ Thiếu file template <code>{$config['template']}</code>. Chạy <code>php build_templates.php</code> để tạo lại.</h2>";
        echo "<a href='index.php'><button>⬅ Quay lại</button></a>";
        exit;
    }

    // 👉 2 ngày của hợp đồng nguyên tắc do người dùng chọn ở form:
    //    - ngày hợp đồng: số hợp đồng lấy tháng/năm của ngày này (01/2026 => 012026)
    //    - ngày ký: điền vào câu "ký ngày ... tháng ... năm ..."
    $hdntHopDongDate = !empty($_POST['hdnt_hd_ngay'])
        ? DateTime::createFromFormat('Y-m-d', $_POST['hdnt_hd_ngay'])
        : null;

    $hdntKyDate = !empty($_POST['hdnt_ngay'])
        ? DateTime::createFromFormat('Y-m-d', $_POST['hdnt_ngay'])
        : null;

    if (!$hdntHopDongDate || !$hdntKyDate) {
        echo "<h2 style='color: red;'>❌ Chưa chọn đủ ngày hợp đồng nguyên tắc và ngày ký.</h2>";
        echo "<a href='index.php'><button>⬅ Quay lại</button></a>";
        exit;
    }

    foreach ($data as $row) {
        $maSoThue = $row['L'] ?? '';
        if (!isset($companies[$maSoThue])) {
            echo "<h2 style='color: red;'>❌ Không tìm thấy công ty có mã số thuế: <strong>$maSoThue</strong></h2>";
            echo "<a href='index.php'><button>⬅ Quay lại</button></a>";
            exit;
        }

        $companyInfo = $companies[$maSoThue];
        $codeName    = $companyInfo['code_name'];

        // 👉 BBGN mang ngày hóa đơn, ĐĐH lùi 10 ngày so với hóa đơn
        $ngayHoaDon = parseExcelDate($row['E'] ?? '');
        if (!$ngayHoaDon) {
            echo "<div style='color: red;'>❌ Bỏ qua 1 dòng: không đọc được ngày ở cột E ("
                . htmlspecialchars((string)($row['E'] ?? '')) . ")</div>";
            continue;
        }

        $ngayDonDatHang = (clone $ngayHoaDon)->modify('-10 days');
        $ngayVanBan     = $isDonDatHang ? $ngayDonDatHang : $ngayHoaDon;

        // 👉 Số đơn đặt hàng viết theo dạng ddmm/yyyy như file mẫu
        $soDonDatHang = $ngayDonDatHang->format('dm/Y');

        $template = new TemplateProcessor($config['template']);
        $template->setValues(array_map('wordValue', [
            // Bên A - bên mua, tra theo mã số thuế ở cột L
            'BenA_Ten'       => $companyInfo['ten'],
            'BenA_DaiDien'   => $companyInfo['daidien'],
            'BenA_ChucVu'    => $companyInfo['chucvu'],
            'BenA_DiaChi'    => $companyInfo['diachi'],
            'BenA_MaSoThue'  => $maSoThue,
            'BenA_DienThoai' => $companyInfo['sdt'],
            'BenA_TaiKhoan'  => $companyInfo['taikhoan'],

            // Hợp đồng nguyên tắc — số mang tháng/năm của ngày hợp đồng
            'HDNT_So'    => $hdntHopDongDate->format('mY') . '/HĐNT PT- '
                . ($companyInfo['ten_viet_tat'] ?? str_replace('_', ' ', $codeName)),
            'HDNT_Ngay'  => $hdntKyDate->format('d'),
            'HDNT_Thang' => $hdntKyDate->format('m'),
            'HDNT_Nam'   => $hdntKyDate->format('Y'),

            // Ngày trên biên bản giao nhận = ngày hóa đơn
            'BBGN_Ngay'  => $ngayHoaDon->format('d'),
            'BBGN_Thang' => $ngayHoaDon->format('m'),
            'BBGN_Nam'   => $ngayHoaDon->format('Y'),

            // Đơn đặt hàng
            'DDH_So'      => $soDonDatHang . '/ ĐĐH',
            'DDH_NgayGon' => $soDonDatHang,
            'DDH_Ngay'    => $ngayDonDatHang->format('d'),
            'DDH_Thang'   => $ngayDonDatHang->format('m'),
            'DDH_Nam'     => $ngayDonDatHang->format('Y'),

            // Hàng hóa và tiền
            'TenHangHoa'        => $row['M'] ?? '',
            'DonViTinh'         => 'Kg',
            'SoLuong'           => formatNumber($row['N'] ?? ''),
            'DonGia'            => formatNumber($row['O'] ?? ''),
            'ThanhTien'         => formatNumber($row['P'] ?? ''),
            'ThueSuat'          => $row['Q'] ?? '',
            'TienThueGTGT'      => formatNumber($row['R'] ?? ''),
            'TongTienThanhToan' => formatNumber($row['S'] ?? ''),
            'TongTienBangChu'   => numberToVietnameseWords(toNumber($row['S'] ?? '')),
        ]));

        if (!isset($companyDirs[$codeName])) {
            $companyDirs[$codeName] = makeExportDir($exportBase, folderName($companyInfo), $exportStamp);
        }
        $companyDir = $companyDirs[$codeName];

        $fileName = "{$config['prefix']}_{$ngayVanBan->format('d_m_Y')}_{$codeName}_{$index}.docx";
        $template->saveAs($companyDir . $fileName);

        $index++;
    }

    $tenTaiLieu = $isDonDatHang ? 'đơn đặt hàng' : 'biên bản giao nhận';
    $ketQuaLink = exportLink($exportLink, $companyDirs);
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
        <p><strong>✅ Thành công!</strong> Đã tạo " . ($index - 1) . " $tenTaiLieu tại thư mục <code>{$ketQuaLink}</code>.</p>
        <div style='margin-top: 15px;'>
            <a href='{$ketQuaLink}' target='_blank' style='
                display: inline-block;
                padding: 10px 20px;
                margin-right: 10px;
                background-color: #28a745;
                color: white;
                text-decoration: none;
                border-radius: 6px;
                font-weight: bold;
            '>📂 Xem tài liệu</a>

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

/**
 * Create Word document using simplified PhpWord approach
 */
function createWordDirectly($data, $outputPath) {
    try {
        $phpWord = new PhpWord();
        $phpWord->setDefaultFontName('Times New Roman');
        $phpWord->setDefaultFontSize(12);
        
        // Tạo hợp đồng mua bán (trang 1-2) với margin tối thiểu
        $section1 = $phpWord->addSection([
            'marginTop' => 400,     // 0.25 inch
            'marginRight' => 430,  // 0.25 inch
            'marginBottom' => 400, // 0.25 inch
            'marginLeft' => 600,   // 0.25 inch
        ]);

        // Header với khoảng cách tối thiểu
        $section1->addText('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', ['bold' => true, 'size' => 14], ['alignment' => 'center']);
        $section1->addText('Độc lập - Tự do - Hạnh phúc', ['bold' => true, 'italic' => true, 'size' => 13], ['alignment' => 'center']);
        $section1->addText('---- o0o ----', ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        $section1->addTextBreak();
        $section1->addText('HỢP ĐỒNG MUA BÁN', ['bold' => true, 'size' => 18], ['alignment' => 'center']);

        // Contract Number
        $contractNumberText = $section1->addTextRun(['alignment' => 'center']);
        $contractNumberText->addText('Số: ', ['bold' => true, 'size' => 13]);
        $contractNumberText->addText($data['NgayThangNamHD'] . '/HĐMB/PT-' . $data['Info_TenCongTy'], ['size' => 13]);

        // Legal Basis
        $section1->addText('     - Căn cứ vào các Bộ luật dân sự Nước Cộng hòa Xã hội Chủ nghĩa Việt Nam', ['size' => 12, 'italic' => true]);
        $section1->addText('     - Căn cứ vào Luật Thương mại 2005 của Nước Cộng hòa Xã hội Chủ nghĩa Việt Nam', ['size' => 12, 'italic' => true]);
        $section1->addText('     - Căn cứ vào nhu cầu và khả năng cung cấp hàng hóa của các bên.', ['size' => 12, 'italic' => true]);

        // Cho riêng King
//        $section1->addText('     - Căn cứ theo luật Thương mại nước Cộng Hòa Xã Hội Chủ Nghĩa Việt Nam được Quốc hội ban hành ngày 27/6/2005, có hiệu lực thi hành từ ngày 01/01/2006;', ['size' => 12, 'italic' => true]);
//        $section1->addText('     - Căn cứ theo bộ luật Dân sự nước Cộng Hòa Xã Hội Chủ Nghĩa Việt Nam được Quốc hội ban hành ngày 27/6/2005, có hiệu lực thi hành từ ngày 01/01/2006;', ['size' => 12, 'italic' => true]);
//        $section1->addText('     - Căn cứ nhu cầu và khả năng của hai bên;', ['size' => 12, 'italic' => true]);

        // Date and Parties Introduction
        $section1->addText('     Hôm nay, ngày    tháng    năm ' . $data['NamHienTai'] . ' chúng tôi gồm:', ['size' => 12, 'italic' => true]);

        // Section I: Seller
        $section1->addText('I.  BÊN BÁN (BÊN A): CÔNG TY TNHH SẢN XUẤT VÀ THƯƠNG MẠI DỊCH VỤ PHÚ THƯ', ['size' => 13, 'bold' => true]);
        $section1->addText('Địa chỉ: Nhà số 1 Thôn Rùa Thượng, Xã Tam Hưng, Thành phố Hà Nội, Việt Nam', ['size' => 13]);
        $section1->addText('Điện thoại: 034.433.2918                                    Fax:', ['size' => 13]);
        $section1->addText('Mã số thuế: 0110430598', ['size' => 13]);
        $representativeText = $section1->addTextRun();
        $representativeText->addText('Đại diện: ', ['size' => 13]);
        $representativeText->addText('Ông Nguyễn Văn Báo', ['bold' => true, 'size' => 13]);
        $representativeText->addText('                       Chức vụ: ', ['size' => 13]);
        $representativeText->addText('Giám đốc', ['bold' => true, 'size' => 13]);
        $section1->addText('Tài khoản: 117002948428 – Tại Ngân hàng Vietinbank – Chi nhánh Thành An', ['size' => 13]);

        // Section II: Buyer
        $section1->addText('II.  BÊN MUA (BÊN B): ' . $data['Info_TenDonViDayDu'], ['size' => 13, 'bold' => true]);
        $section1->addText('Địa chỉ: ' . $data['Info_DiaChi'], ['size' => 13]);
        $section1->addText('Điện thoại: ' . $data['Info_DienThoai'] . '                             Fax: ' . $data['Info_Fax'], ['size' => 13]);
        $section1->addText('Mã số thuế: ' . $data['Info_MaSoThue'], ['size' => 13]);
        $representativeBText = $section1->addTextRun();
        $representativeBText->addText('Đại diện: ', ['size' => 13]);
        $representativeBText->addText($data['Info_DaiDien'], ['bold' => true, 'size' => 13]);
        $representativeBText->addText('                            Chức vụ: ', ['size' => 13]);
        $representativeBText->addText($data['Info_ChucVu'], ['bold' => true, 'size' => 13]);
        $section1->addText('Tài khoản: ' . $data['Info_TaiKhoan'], ['size' => 13]);

        // Agreement Statement
        $section1->addText('   Sau khi thỏa thuận, bên B đồng ý mua, Bên A đồng ý bán hàng hóa theo các điều kiện sau đây:', ['size' => 13]);

        // Article 1
        $article1Text = $section1->addTextRun();
        $article1Text->addText('Điều 1', ['bold' => true, 'size' => 13, 'underline' => 'single']);
        $article1Text->addText(': Hàng hóa, chất lượng và số lượng.', ['bold' => true, 'size' => 13]);
        $item11Text = $section1->addTextRun();
        $item11Text->addText('1.1. ', ['bold' => true, 'size' => 13]);
        $item11Text->addText('Hàng hóa và số lượng: Bên B đồng ý mua của bên A mặt hàng ', ['size' => 13]);
        $item11Text->addText($data['TenHangHoaDichVu'], ['bold' => true, 'size' => 13]);
        $item11Text->addText(' do bên A bán với số lượng chi tiết ở điều 2.1 của hợp đồng.', ['size' => 13]);
        $item12Text = $section1->addTextRun();
        $item12Text->addText('1.2. ', ['bold' => true, 'size' => 13]);
        $item12Text->addText('Chất lượng: Hàng hóa mới 100%.', ['size' => 13]);

        // Article 2
        $article2Text = $section1->addTextRun();
        $article2Text->addText('Điều 2', ['bold' => true, 'size' => 13, 'underline' => 'single']);
        $article2Text->addText(': Giá cả, phương thức thanh toán.', ['bold' => true, 'size' => 13]);
        $item21Text = $section1->addTextRun();
        $item21Text->addText('2.1. Giá cả: ', ['bold' => true, 'size' => 13]);
        $item21Text->addText('giá trị của hàng hóa nêu tại Điều 1 của hợp đồng này được tính theo bảng giá như sau:', ['size' => 13]);

        // Chuyển bảng sang trang mới
        $section1->addTextBreak();
        $section1->addTextBreak();
        $section1->addTextBreak();

        // Price Table với định dạng tối thiểu - tăng cellMargin để có khoảng cách
        $tableStyle = [
            'borderSize' => 6,
            'borderColor' => '000000',
            'cellMargin' => 200, // Tăng từ 50 lên 200 để có khoảng cách
            'cellSpacing' => 0,
        ];
        $table = $section1->addTable($tableStyle);

        // Header row với định dạng tối thiểu - căn giữa trái phải và trên dưới, cellMargin = 200, size 13
        $headerRow = $table->addRow();
        $headerRow->addCell(1000, ['bgColor' => 'FFFFFF', 'valign' => 'center', 'cellMargin' => 200])->addText('STT', ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        $headerRow->addCell(3500, ['bgColor' => 'FFFFFF', 'valign' => 'center', 'cellMargin' => 200])->addText('LOẠI HÀNG', ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        $headerRow->addCell(1000, ['bgColor' => 'FFFFFF', 'valign' => 'center', 'cellMargin' => 200])->addText('ĐVT', ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        $headerRow->addCell(1500, ['bgColor' => 'FFFFFF', 'valign' => 'center', 'cellMargin' => 200])->addText('S.LƯỢNG', ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        $headerRow->addCell(2000, ['bgColor' => 'FFFFFF', 'valign' => 'center', 'cellMargin' => 200])->addText('ĐƠN GIÁ', ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        $headerRow->addCell(2000, ['bgColor' => 'FFFFFF', 'valign' => 'center', 'cellMargin' => 200])->addText('THÀNH TIỀN', ['bold' => true, 'size' => 13], ['alignment' => 'center']);

        // Data row - căn giữa trái phải và trên dưới, cellMargin = 150, size 13
        $dataRow = $table->addRow();
        $dataRow->addCell(1000, ['valign' => 'center', 'cellMargin' => 150])->addText('1', ['size' => 13], ['alignment' => 'center']);
        $dataRow->addCell(3500, ['valign' => 'center', 'cellMargin' => 150])->addText($data['TenHangHoaDichVu'], ['size' => 13], ['alignment' => 'center']);
        $dataRow->addCell(1000, ['valign' => 'center', 'cellMargin' => 150])->addText('Kg', ['size' => 13], ['alignment' => 'center']);
        $dataRow->addCell(1500, ['valign' => 'center', 'cellMargin' => 150])->addText($data['SoLuong'], ['size' => 13], ['alignment' => 'center']);
        $dataRow->addCell(2000, ['valign' => 'center', 'cellMargin' => 150])->addText($data['DonGia'], ['size' => 13], ['alignment' => 'center']);
        $dataRow->addCell(2000, ['valign' => 'center', 'cellMargin' => 150])->addText($data['ThanhTien'], ['size' => 13], ['alignment' => 'center']);

        // Summary rows với định dạng tối thiểu - merge 5 cột đầu thành 1 cột, căn giữa trái phải và trên dưới, cellMargin = 150, size 13
        $summaryRow1 = $table->addRow();
        $summaryRow1->addCell(1000, ['gridSpan' => 5, 'valign' => 'center', 'cellMargin' => 150])->addText('Tổng cộng tiền hàng', ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        $summaryRow1->addCell(2000, ['valign' => 'center', 'cellMargin' => 150])->addText($data['ThanhTien'], ['bold' => true, 'size' => 13], ['alignment' => 'center']);

        $summaryRow2 = $table->addRow();
        $summaryRow2->addCell(1000, ['gridSpan' => 5, 'valign' => 'center', 'cellMargin' => 150])->addText('Thuế VAT 10%', ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        $summaryRow2->addCell(2000, ['valign' => 'center', 'cellMargin' => 150])->addText($data['TienThueGTGT'], ['bold' => true, 'size' => 13], ['alignment' => 'center']);

        $summaryRow3 = $table->addRow();
        $summaryRow3->addCell(1000, ['gridSpan' => 5, 'valign' => 'center', 'cellMargin' => 150])->addText('Tổng cộng', ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        $summaryRow3->addCell(2000, ['valign' => 'center', 'cellMargin' => 150])->addText($data['TongTienThanhToan'], ['bold' => true, 'size' => 13], ['alignment' => 'center']);

        $bangChuText = $section1->addTextRun();
        $bangChuText->addText('Bằng chữ: ', ['bold' => true, 'size' => 13]);
        $bangChuText->addText($data['TongTienThanhToan_BangChu'], ['bold' => true, 'italic' => true, 'size' => 13]);
        $section1->addText('Giá trị hợp đồng được phép dung sai tối đa +/-10%', ['bold' => true, 'italic' => true, 'size' => 13]);

        // King thì comment lại
//        $section1->addText('Giá trị hợp đồng được phép dung sai tối đa +/-10%', ['bold' => true, 'italic' => true, 'size' => 13]);

        $section1->addText('2.2. Phương thức thanh toán:', ['bold' => true, 'size' => 13]);
        $section1->addText('     - Hình thức thanh toán: Thanh toán trước hoặc sau không quá 60 ngày khi bên A giao đủ hàng và chứng từ cho bên B, bằng tiền mặt hoặc chuyển khoản.', ['size' => 13]);
        $section1->addText('     - Chứng từ thanh toán bao gồm: Hoá đơn GTGT, Biên bản giao nhận hàng hóa.', ['size' => 13]);

        // Article 3
        $article3Text = $section1->addTextRun();
        $article3Text->addText('Điều 3', ['bold' => true, 'size' => 13, 'underline' => 'single']);
        $article3Text->addText(': Địa điểm và phương thức giao nhận hàng hóa', ['bold' => true, 'size' => 13]);
        $item31Text = $section1->addTextRun();
        $item31Text->addText('3.1. Địa điểm giao nhận: ', ['bold' => true, 'size' => 13]);
        $item31Text->addText('Tại xưởng bán bên A', ['size' => 13]);

        $item32Text = $section1->addTextRun();
        $item32Text->addText('3.2. Phương thức giao nhận: ', ['bold' => true, 'size' => 13]);
        $item32Text->addText('Giao hàng qua cân thực tế và lập Biên bản giao nhận hàng hóa giữa hai bên.Khối lượng trên biên bản được xác định làm căn cứ thanh toán. Nếu phát hiện hàng thiếu hoặc không đúng tiêu chuẩn chất lượng thì lập biên bản tại chỗ và yêu cầu bên A xác nhận. Trường hợp hàng không đạt chất lượng như quy định tại điều 2, bên B có quyền từ chối không nhận hàng.', ['size' => 13]);

        $item33Text = $section1->addTextRun();
        $item33Text->addText('3.3. Thời hạn giao hàng: ', ['bold' => true, 'size' => 13]);
        $item33Text->addText('Hàng giao theo yêu cầu bên mua', ['size' => 13]);

        // Article 4
        $article4Text = $section1->addTextRun();
        $article4Text->addText('Điều 4', ['bold' => true, 'size' => 13, 'underline' => 'single']);
        $article4Text->addText(': Quyền và nghĩa vụ của các bên', ['bold' => true, 'size' => 13]);
        $section1->addText('4.1. Quyền và nghĩa vụ của Bên B:', ['bold' => true, 'size' => 13]);
        $section1->addText('- Kiểm tra số lượng, chất lượng hàng hóa nêu tại Điều 1 của hợp đồng này.', ['size' => 13]);
        $section1->addText('- Chuẩn bị nhân lực để tiếp nhận hàng khi có thông báo giao hàng của Bên A', ['size' => 13]);
        $section1->addText('- Đảm bảo việc thanh toán theo đúng giá trị và phương thức thanh toán đã nêu tại Điều 2 của Hợp đồng này.', ['size' => 13]);
        $section1->addText('- Bảo quản sử dụng hàng hóa theo đúng hướng dẫn của Nhà sản xuất.', ['size' => 13]);
        $section1->addText('- Có quyền khiếu nại về hàng hóa bên A giao trong thời hạn 02 ngày về số lượng, chủng loại và trong thời hạn 15 ngày về chất lượng kể từ ngày bên A giao hàng.', ['size' => 13]);
        $section1->addText('4.2. Quyền và nghĩa vụ của Bên A:', ['bold' => true, 'size' => 13]);
        $section1->addText('- Yêu cầu Bên B thực hiện đúng và hai bên đã cam kết.', ['size' => 13]);
        $section1->addText('- Đảm bảo cung cấp đúng và đầy đủ số lượng, chất lượng hàng hóa nêu tại Điều 1 của Hợp đồng', ['size' => 13]);
        $section1->addText('- Giải quyết các khiếu nại của bên B trong thời hạn 05 ngày kể từ ngày nhận được thư khiếu nại của bên A.', ['size' => 13]);

        // Article 5
        $article5Text = $section1->addTextRun();
        $article5Text->addText('Điều 5', ['bold' => true, 'size' => 13, 'underline' => 'single']);
        $article5Text->addText(': Điều khoản thi hành', ['bold' => true, 'size' => 13]);
        $item51Text = $section1->addTextRun();
        $item51Text->addText('5.1. ', ['bold' => true, 'size' => 13]);
        $item51Text->addText('Hai bên cam kết thực hiện đúng và đầy đủ các điều khoản đã cam kết trong Hợp đồng.', ['size' => 13]);

        $item52Text = $section1->addTextRun();
        $item52Text->addText('5.2. ', ['bold' => true, 'size' => 13]);
        $item52Text->addText('Trong quá trình thực hiện, nếu có phát sinh những bất lợi và sự kiện mới thì hai bên cùng nhau bàn bạc, nhất trí thỏa thuận về việc sửa đổi, bổ sung Hợp đồng. Mọi sửa đổi và bổ sung của Hợp đồng chỉ có giá trị khi được hai bên thoả thuận và lập thành văn bản.', ['size' => 13]);

        $item53Text = $section1->addTextRun();
        $item53Text->addText('5.3. ', ['bold' => true, 'size' => 13]);
        $item53Text->addText('Trường hợp có phát sinh tranh chấp thì các bên cùng tiến hành thương lượng giải quyết trên cơ sở tôn trọng quyền lợi của nhau. Nếu các bên không giải quyết được thì một trong hai bên có quyền yêu cầu Tòa án có thẩm quyền giải quyết.theo quy định của pháp luật. Phán quyết của Tòa án là căn cứ để hai bên thực hiện. Nếu bên nào sai thì hoàn toàn chịu mọi phí tổn Tòa án.', ['size' => 13]);

        $item55Text = $section1->addTextRun();
        $item55Text->addText('5.5. ', ['bold' => true, 'size' => 13]);
        $item55Text->addText('Hợp đồng được lập thành 02 bản, mỗi bên giữ 01 bản, các bản có giá trị pháp lý như nhau.', ['size' => 13]);

        $item56Text = $section1->addTextRun();
        $item56Text->addText('5.6. ', ['bold' => true, 'size' => 13]);
        $item56Text->addText('Hợp đồng có hiệu lực kể từ ngày ký và tự động hết hiệu lực khi hai bên hoàn thành nghĩa vụ của mình. Trong thời hạn 05 ngày kể từ ngày bên A hoàn thành nghĩa vụ thanh toán cho bên B, hợp đồng coi như được thanh lý.', ['size' => 13]);
        $section1->addTextBreak();

        // Signatures
        $section1->addText('ĐẠI DIỆN BÊN A            	                                    	              ĐẠI DIỆN BÊN B', ['bold' => true, 'size' => 14], ['alignment' => 'center']);


        // Biên bản giao nhận

        // Tạo biên bản giao nhận hàng hóa (trang 3-4) với margin tối thiểu
        $section2 = $phpWord->addSection([
            'marginTop' => 400,     // ~0.28 inch
            'marginRight' => 430,  // ~0.28 inch
            'marginBottom' => 400, // ~0.28 inch
            'marginLeft' => 600,   // ~0.42 inch
        ]);
        
        // Header với khoảng cách tối thiểu
        $section2->addText('CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM', ['bold' => true, 'size' => 14], ['alignment' => 'center']);
        $section2->addText('Độc lập - Tự do - Hạnh phúc', ['bold' => true, 'italic' => true, 'size' => 13], ['alignment' => 'center']);
        $section2->addText('~~~~~o0o~~~~~', ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        $section2->addText('BIÊN BẢN GIAO NHẬN HÀNG HOÁ KIÊM PHIẾU XUẤT KHO', ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        
        $section2->addText('Hôm nay, ngày    tháng   năm ' . $data['NamHienTai'] . ' tại Công ty TNHH sản xuất và thương mại dịch vụ Phú Thư , chúng tôi gồm có:', ['italic' => true, 'size' => 13]);
        
        // Section 1: Seller
        $section2->addText('1. BÊN A (BÊN BÁN) : CÔNG TY TNHH SẢN XUẤT VÀ THƯƠNG MẠI DỊCH VỤ PHÚ THƯ', ['bold' => true, 'size' => 13]);
        $section2->addText('Địa chỉ: Nhà số 1 thôn Rùa Thượng, Xã Tam Hưng, Thành phố Hà Nội', ['size' => 13]);
        $section2->addText('Điện thoại: 034.433.2918            	Fax:', ['size' => 13]);
        $section2->addText('Mã số thuế: 0110430598', ['bold' => true, 'size' => 13]);
        $section2->addText('Đại diện: Bà Lý Ngọc Mai            	Chức vụ: Kế toán', ['bold' => true, 'size' => 13]);
        
        // Section 2: Buyer
        $section2->addText('2. BÊN B (BÊN MUA) : ' . $data['Info_TenDonViDayDu'], ['bold' => true, 'size' => 13]);
        $section2->addText('Địa chỉ: ' . $data['Info_DiaChi'], ['size' => 13]);
        $section2->addText('Điện thoại: ' . $data['Info_DienThoai'] . '	Fax: ' . $data['Info_Fax'], ['size' => 13]);
        $section2->addText('Mã số thuế: ' . $data['Info_MaSoThue'], ['bold' => true, 'size' => 13]);
        $section2->addText('Đại diện: …………………….	Chức vụ: ………………..', ['bold' => true, 'size' => 13]);
        
        $section2->addText('Hai bên thống nhất lập biên bản giao nhận hàng hoá với các nội dung sau:', ['italic' => true, 'size' => 13]);
        
        // Delivery Table với định dạng tối thiểu - giống hệt table1
        $table2Style = [
            'borderSize' => 6,
            'borderColor' => '000000',
            'cellMargin' => 200, // Giống table1
            'cellSpacing' => 0,
        ];
        $table2 = $section2->addTable($table2Style);
        
        // Header row với định dạng tối thiểu - căn giữa trái phải và trên dưới, cellMargin = 200, size 13
        $headerRow2 = $table2->addRow();
        $headerRow2->addCell(1000, ['bgColor' => 'FFFFFF', 'valign' => 'center', 'cellMargin' => 200])->addText('TT', ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        $headerRow2->addCell(3500, ['bgColor' => 'FFFFFF', 'valign' => 'center', 'cellMargin' => 200])->addText('Tên hàng hóa', ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        $headerRow2->addCell(1500, ['bgColor' => 'FFFFFF', 'valign' => 'center', 'cellMargin' => 200])->addText('Đơn vị tính', ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        $headerRow2->addCell(1500, ['bgColor' => 'FFFFFF', 'valign' => 'center', 'cellMargin' => 200])->addText('Số lượng', ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        $headerRow2->addCell(2000, ['bgColor' => 'FFFFFF', 'valign' => 'center', 'cellMargin' => 200])->addText('Đơn giá', ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        $headerRow2->addCell(2000, ['bgColor' => 'FFFFFF', 'valign' => 'center', 'cellMargin' => 200])->addText('Thành tiền', ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        
        // Data row - căn giữa trái phải và trên dưới, cellMargin = 150, size 13
        $dataRow2 = $table2->addRow();
        $dataRow2->addCell(1000, ['valign' => 'center', 'cellMargin' => 150])->addText('1', ['size' => 13], ['alignment' => 'center']);
        $dataRow2->addCell(3500, ['valign' => 'center', 'cellMargin' => 150])->addText($data['TenHangHoaDichVu'], ['size' => 13], ['alignment' => 'center']);
        $dataRow2->addCell(1500, ['valign' => 'center', 'cellMargin' => 150])->addText('Kg', ['size' => 13], ['alignment' => 'center']);
        $dataRow2->addCell(1500, ['valign' => 'center', 'cellMargin' => 150])->addText($data['SoLuong'], ['size' => 13], ['alignment' => 'center']);
        $dataRow2->addCell(2000, ['valign' => 'center', 'cellMargin' => 150])->addText($data['DonGia'], ['size' => 13], ['alignment' => 'center']);
        $dataRow2->addCell(2000, ['valign' => 'center', 'cellMargin' => 150])->addText($data['ThanhTien'], ['size' => 13], ['alignment' => 'center']);
        
        // Summary rows với định dạng tối thiểu - merge 5 cột đầu thành 1 cột, căn giữa trái phải và trên dưới, cellMargin = 150, size 13
        $summaryRow2_1 = $table2->addRow();
        $summaryRow2_1->addCell(9500, ['gridSpan' => 5, 'valign' => 'center', 'cellMargin' => 150])->addText('Giá trước thuế', ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        $summaryRow2_1->addCell(2000, ['valign' => 'center', 'cellMargin' => 150])->addText($data['ThanhTien'], ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        
        $summaryRow2_2 = $table2->addRow();
        $summaryRow2_2->addCell(9500, ['gridSpan' => 5, 'valign' => 'center', 'cellMargin' => 150])->addText('Thuế GTGT 10%', ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        $summaryRow2_2->addCell(2000, ['valign' => 'center', 'cellMargin' => 150])->addText($data['TienThueGTGT'], ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        
        $summaryRow2_3 = $table2->addRow();
        $summaryRow2_3->addCell(9500, ['gridSpan' => 5, 'valign' => 'center', 'cellMargin' => 150])->addText('Tổng cộng', ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        $summaryRow2_3->addCell(2000, ['valign' => 'center', 'cellMargin' => 150])->addText($data['TongTienThanhToan'], ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        
        $bangChuText2 = $section2->addTextRun();
        $bangChuText2->addText('Số tiền bằng chữ: ', ['bold' => true, 'size' => 13]);
        $bangChuText2->addText($data['TongTienThanhToan_BangChu'], ['bold' => true, 'italic' => true, 'size' => 13]);
        
        $section2->addText('Biên bản này được lập thành 02 bản có giá trị pháp lý như nhau, bên A giữ 01 bản, bên B giữ 01 bản .', ['size' => 13]);
        
        // Signatures
        $section2->addText('ĐẠI DIỆN BÊN BÁN                                                                             ĐẠI DIỆN BÊN MUA', ['bold' => true, 'size' => 13], ['alignment' => 'center']);
        
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
            // Đảm bảo encoding UTF-8
            $cleanValue = mb_convert_encoding((string)$value, 'UTF-8', 'UTF-8');
            // Loại bỏ ký tự không in được (giữ nguyên dấu &)
//            $cleanValue = preg_replace('/[\x00-\x1F\x7F-\x9F]/', '', $cleanValue);
            return $cleanValue;
        };
        
        
        // Map dữ liệu theo đúng template DOCX gốc - đầy đủ các field
        $tongTienNumeric = (float)str_replace(',', '', $safeValue($row['S']));
        $documentData = [
            'Info_TenDonViDayDu' => $safeValue($companyInfo['ten']), // Thay thế & bằng "và"
            'Info_MaSoThue' => $safeValue($maSoThue),
            'Info_DiaChi' => $safeValue($companyInfo['diachi']),
            'Info_DienThoai' => $safeValue($companyInfo['sdt']),
            'Info_Fax' => $safeValue($companyInfo['fax']),
            'Info_DaiDien' => $safeValue($companyInfo['daidien']),
            'Info_ChucVu' => $safeValue($companyInfo['chucvu']),
            'Info_TaiKhoan' => $safeValue($companyInfo['taikhoan']),
            'Info_TenCongTy' => $safeValue($companyInfo['code']),
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
        if (!isset($companyDirs[$codeName])) {
            $companyDirs[$codeName] = makeExportDir($exportBase, folderName($companyInfo), $exportStamp);
        }
        $companyDir = $companyDirs[$codeName];

        $fileName = "HĐMB_{$originalYear}_{$codeName}_{$index}_" . time() . ".docx";
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

    $ketQuaLink = exportLink($exportLink, $companyDirs);
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
        <p><strong>✅ Thành công!</strong> Đã tạo " . ($index - 1) . " hợp đồng từ HTML template tại thư mục <code>{$ketQuaLink}</code>.</p>
        <div style='margin-top: 15px;'>
            <a href='{$ketQuaLink}' target='_blank' style='
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
    if (!isset($companyDirs[$codeName])) {
        $companyDirs[$codeName] = makeExportDir($exportBase, folderName($companyInfo), $exportStamp);
    }
    $companyDir = $companyDirs[$codeName];

    $fileName = "HĐMB_{$originalYear}_{$codeName}_{$index}.docx";
    $savePath = $companyDir . $fileName;
    $template->saveAs($savePath);

    $index++;
}


$ketQuaLink = exportLink($exportLink, $companyDirs);
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
    <p><strong>✅ Thành công!</strong> Đã tạo " . ($index - 1) . " hợp đồng tại thư mục <code>{$ketQuaLink}</code>.</p>
    <div style='margin-top: 15px;'>
        <a href='{$ketQuaLink}' target='_blank' style='
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
