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

        <select name="doc_type" required style="margin-bottom: 1rem; padding: 0.6rem; width: 100%;">
            <!-- <option value="HĐMB">Hợp Đồng Mua Bán (Word - DOCX Template)</option> -->
            <option value="HĐMB_HTML">Hợp Đồng Mua Bán (Word - HTML Template)</option>
            <option value="BBGN">Biên Bản Giao Nhận (Excel)</option>
        </select>

        <button type="submit">Tạo Tài Liệu</button>
    </form>

</div>
</body>
</html>
