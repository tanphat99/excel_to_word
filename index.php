<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <title>Upload Excel và Xuất Word</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600&display=swap" rel="stylesheet">
    <style>
        body {
            font-family: 'Inter', sans-serif;
            background: #f7f9fc;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            margin: 0;
        }

        .container {
            background: #ffffff;
            padding: 2rem 2.5rem;
            border-radius: 12px;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.08);
            text-align: center;
            width: 100%;
            max-width: 400px;
        }

        h2 {
            margin-bottom: 1.5rem;
            color: #333333;
        }

        input[type="file"] {
            display: block;
            margin: 0 auto 1.5rem auto;
            padding: 0.6rem;
            border: 2px dashed #c3c6cf;
            border-radius: 8px;
            width: 100%;
            background-color: #fafafa;
            cursor: pointer;
            transition: border-color 0.3s;
        }

        input[type="file"]:hover {
            border-color: #007bff;
        }

        .doc-row {
            display: flex;
            gap: 0.5rem;
        }

        .doc-row select {
            flex: 1;
            min-width: 0;
        }

        .doc-row select,
        .doc-row input[type="date"] {
            padding: 0.6rem;
            border: 1px solid #c3c6cf;
            border-radius: 8px;
            background-color: #fff;
            font-family: inherit;
            font-size: 0.95rem;
            color: #333;
        }

        .doc-row input[type="date"]:disabled {
            background-color: #f1f2f5;
            color: #9a9ea8;
        }

        .hint {
            margin: 0.5rem 0 1.5rem;
            font-size: 0.8rem;
            color: #7a7f8a;
            text-align: left;
            line-height: 1.4;
        }

        .hint.muted {
            opacity: 0.5;
        }

        button {
            background-color: #007bff;
            color: white;
            border: none;
            padding: 0.75rem 1.5rem;
            font-size: 1rem;
            font-weight: 600;
            border-radius: 8px;
            cursor: pointer;
            transition: background-color 0.3s;
        }

        button:hover {
            background-color: #0056b3;
        }

        @media (max-width: 480px) {
            .container {
                padding: 1.5rem;
            }

            button {
                width: 100%;
            }
        }
    </style>
</head>
<body>
<div class="container">
    <h2>Upload Excel</h2>
    <form action="process.php" method="post" enctype="multipart/form-data">
        <input type="file" name="excel_file" accept=".xlsx,.xls" required>

        <div class="doc-row">
            <select name="doc_type" id="doc_type" required>
                <!-- <option value="HĐMB">Hợp Đồng Mua Bán (Word - DOCX Template)</option> -->
                <option value="HĐMB_HTML">Hợp Đồng Mua Bán (Word - HTML Template)</option>
                <option value="BBGN">Biên Bản Giao Nhận (Excel)</option>
                <option value="BBGN_WORD">Biên Bản Giao Nhận Hàng Hóa (Word)</option>
                <option value="DDH">Đơn Đặt Hàng (Word)</option>
            </select>

            <input type="date" name="hdnt_ngay" id="hdnt_ngay" title="Ngày ký hợp đồng nguyên tắc">
        </div>

        <p class="hint" id="hdnt_hint">Ngày ký hợp đồng nguyên tắc — dùng cho Biên Bản Giao Nhận (Word) và Đơn Đặt Hàng.</p>

        <button type="submit">Tạo Tài Liệu</button>
    </form>

</div>

<script>
    // Ngày ký hợp đồng nguyên tắc chỉ dùng cho 2 template Word mới
    (function () {
        var docType = document.getElementById('doc_type');
        var ngayKy = document.getElementById('hdnt_ngay');
        var hint = document.getElementById('hdnt_hint');

        function toggle() {
            var canDung = docType.value === 'BBGN_WORD' || docType.value === 'DDH';
            ngayKy.disabled = !canDung;
            hint.classList.toggle('muted', !canDung);
        }

        docType.addEventListener('change', toggle);
        toggle();
    })();
</script>
</body>
</html>
