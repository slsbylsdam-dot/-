<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>النظام المحاسبي الشامل - مخاريج البئر والمشاريع المشتركة</title>
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap" rel="stylesheet">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js"></script>
    <script>
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
    </script>
    <style>
        * { box-sizing: border-box; font-family: 'Cairo', sans-serif; margin: 0; padding: 0; }
        body { background-color: #f4f7f6; color: #333; padding: 10px; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 15px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); }
        h2 { text-align: center; color: #0b5ed7; font-size: 22px; margin-bottom: 5px; }
        .subtitle { text-align: center; font-size: 13px; color: #6c757d; margin-bottom: 20px; line-height: 1.5; }
        .cards-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px; margin-bottom: 20px; }
        .card { background: #e7f1ff; border-right: 5px solid #0d6efd; padding: 12px; border-radius: 8px; }
        .card.partner-card { background: #e8f5e9; border-right-color: #198754; }
        .card h3 { font-size: 13px; color: #495057; margin-bottom: 5px; }
        .card .value { font-size: 15px; font-weight: bold; color: #0b5ed7; }
        .settlement-card { background: #fff; border: 1px solid #dee2e6; border-right: 5px solid #0d6efd; padding: 15px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.02); }
        .form-container { background: #f8f9fa; padding: 15px; border-radius: 8px; margin-bottom: 20px; border: 1px solid #dee2e6; }
        .form-container h3 { margin-bottom: 10px; font-size: 15px; color: #333; }
        .form-group { display: grid; grid-template-columns: 1fr; gap: 10px; }
        @media(min-width: 768px) { .form-group { grid-template-columns: 1fr 2fr 1fr 1fr auto; } }
        .form-group-payment { display: grid; grid-template-columns: 1fr; gap: 10px; }
        @media(min-width: 768px) { .form-group-payment { grid-template-columns: 1fr 2fr 1fr 1.5fr auto; } }
        input, select, button { padding: 10px; font-size: 14px; border: 1px solid #ced4da; border-radius: 6px; width: 100%; }
        button { background-color: #198754; color: white; border: none; cursor: pointer; font-weight: bold; transition: 0.2s; }
        button:hover { background-color: #157347; }
        .action-bar { display: flex; gap: 8px; margin-bottom: 20px; flex-direction: column; }
        @media(min-width: 576px) { .action-bar { flex-direction: row; flex-wrap: wrap; } }
        .action-bar button, .action-bar label { flex: 1; min-width: 110px; font-size: 12px; padding: 10px; text-align: center; border-radius: 6px; cursor: pointer; font-weight: bold; color: white; display: inline-flex; align-items: center; justify-content: center; }
        .btn-pdf { background-color: #dc3545; }
        .btn-whatsapp { background-color: #25d366; }
        .btn-setup { background-color: #6f42c1; }
        .btn-losses { background-color: #fd7e14; }
        .btn-import { background-color: #0d6efd; }
        .btn-import-word { background-color: #0d6efd; }
        .btn-import-pdf { background-color: #d63384; }
        .btn-new-account { background-color: #20c997; }
        .btn-general-stmt { background-color: #0b5ed7; }
        .btn-payment { background-color: #0dcaf0; color: #000; }
        .btn-share-whatsapp { background-color: #25d366; color: white; padding: 8px 12px; font-size: 13px; border-radius: 5px; margin-top: 10px; width: 100%; display: flex; align-items: center; justify-content: center; gap: 5px; }
        .table-responsive { width: 100%; overflow-x: auto; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; min-width: 600px; }
        th, td { padding: 10px 8px; text-align: right; border-bottom: 1px solid #dee2e6; font-size: 13px; }
        th { background-color: #f8f9fa; color: #495057; font-weight: 600; }
        .action-buttons-cell { display: flex; gap: 5px; }
        .btn-action { padding: 6px 10px; font-size: 12px; border-radius: 4px; cursor: pointer; border: none; color: #fff; }
        .edit-btn { background-color: #ffc107; color: #000; }
        .delete-btn { background-color: #dc3545; }
        .badge { background: #e2e3e5; padding: 4px 8px; border-radius: 4px; font-size: 11px; font-weight: bold; display: inline-block; }
        .badge-partner { background: #d1e7dd; color: #0f5132; }
        .badge-payment { background: #cff4fc; color: #055160; }
        .pdf-hidden { display: none !important; }
        .modal-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 9999; justify-content: center; align-items: center; padding: 15px; }
        .modal-content { background: #fff; width: 100%; max-width: 850px; max-height: 90vh; overflow-y: auto; padding: 20px; border-radius: 12px; box-shadow: 0 5px 20px rgba(0,0,0,0.2); position: relative; }
        .close-modal { position: absolute; top: 15px; left: 15px; background: #dc3545; color: white; border: none; width: 30px; height: 30px; border-radius: 50%; font-weight: bold; cursor: pointer; display: flex; align-items: center; justify-content: center; }
        .partner-input-row { display: flex; gap: 10px; margin-bottom: 10px; align-items: center; }
    </style>
</head>
<body>

<div class="container" id="pdfContent">
    <div style="background: #e9ecef; padding: 10px; border-radius: 8px; margin-bottom: 15px; display: flex; gap: 10px; align-items: center; flex-wrap: wrap;">
        <label style="font-weight: bold; font-size: 13px; color: #495057;">اختر كشف الحساب:</label>
        <select id="projectSelector" onchange="switchProject(this.value)" style="flex: 1; min-width: 200px; font-weight: bold; color: #0b5ed7;"></select>
    </div>

    <h2 id="displayProjectName">حسابات مخاريج البئر محلي</h2>
    <div class="subtitle" id="displayProjectDesc">تخص البئر المشتركة بين: أولاد عبده أحمد حسين، أولاد صالح أحمد حسين، وأولاد حسين أحمد حسين</div>

    <div class="action-bar" id="actionBar">
        <button class="btn-new-account" type="button" onclick="openNewAccountModal()">📁 مشروع جديد</button>
        <button class="btn-setup" type="button" onclick="openSetupModal()">⚙️ إعدادات الشركاء</button>
        <button class="btn-general-stmt" type="button" onclick="openGeneralModal()">📋 الكشف العام</button>
        <label class="btn-import">
            📥 إكسل
            <input type="file" id="excelFileImport" accept=".xlsx, .xls, .csv" style="display: none;" onchange="importExcelData(event)">
        </label>
        <label class="btn-import-word">
            📝 وورد
            <input type="file" id="wordFileImport" accept=".docx" style="display: none;" onchange="importWordData(event)">
        </label>
        <label class="btn-import-pdf">
            📄 بي دي إف
            <input type="file" id="pdfFileImport" accept=".pdf" style="display: none;" onchange="importPdfData(event)">
        </label>
        <button class="btn-pdf" type="button" onclick="exportPDF()">📄 تصدير PDF</button>
        <button class="btn-whatsapp" type="button" onclick="downloadPdfAndOpenWhatsApp()">💬 واتساب</button>
        <button class="btn-losses" type="button" onclick="openLossesModal()">📉 الخسائر</button>
    </div>

    <div class="form-container" id="formContainer">
        <h3 id="formTitle">➕ إضافة مصروف أو تكلفة جديدة</h3>
        <input type="hidden" id="editIndex" value="-1">
        <div class="form-group">
            <input type="date" id="dateInput">
            <input type="text" id="descInput" placeholder="التفاصيل...">
            <input type="number" id="amountInput" placeholder="المبلغ">
            <div>
                <select id="payerSelect">
                    <option value="صندوق البئر (مشترك)">صندوق البئر (مشترك)</option>
                </select>
            </div>
            <button type="button" id="saveBtn" onclick="saveExpense()">إضافة القيد</button>
        </div>
    </div>

    <div class="form-container" style="background-color: #e3f2fd; border-color: #90caf9;" id="paymentFormContainer">
        <h3 id="paymentFormTitle" style="color: #0d6efd;">💵 تسجيل دفعة نقدية / تسديد من شريك</h3>
        <input type="hidden" id="editPaymentIndex" value="-1">
        <div class="form-group-payment">
            <input type="date" id="payDateInput">
            <input type="text" id="payDescInput" placeholder="بيان الدفعة النقدية">
            <input type="number" id="payAmountInput" placeholder="المبلغ المسدد">
            <div>
                <select id="payPartnerSelect"></select>
            </div>
            <button type="button" class="btn-payment" onclick="savePartnerPayment()" id="savePaymentBtn">تسجيل الدفعة</button>
        </div>
    </div>

    <h3 style="margin-bottom: 10px; color: #0b5ed7; font-size: 16px;">📋 سجل العمليات والمصاريف والتكاليف</h3>
    <div class="table-responsive">
        <table>
            <thead>
                <tr>
                    <th>التاريخ</th>
                    <th>التفاصيل</th>
                    <th>المبلغ</th>
                    <th>الجهة الدافاة</th>
                    <th id="actionHeader">إجراءات</th>
                </tr>
            </thead>
            <tbody id="expenseTableBody"></tbody>
        </table>
    </div>

    <h3 style="margin-top: 25px; margin-bottom: 10px; color: #0d6efd; font-size: 16px;">💳 سجل الدفعات النقدية والتسديدات الخاصة بالشركاء</h3>
    <div class="table-responsive">
        <table>
            <thead>
                <tr>
                    <th>التاريخ</th>
                    <th>البيان</th>
                    <th>المبلغ المسدد</th>
                    <th>الشريك المسدد</th>
                    <th id="paymentActionHeader">إجراءات</th>
                </tr>
            </thead>
            <tbody id="paymentTableBody"></tbody>
        </table>
    </div>

    <hr style="margin: 25px 0; border: 0; border-top: 1px solid #dee2e6;">

    <h3 style="margin-bottom: 12px; color: #0b5ed7; font-size: 16px;">📊 الملخصات العامة وحصص الشركاء</h3>
    <div class="cards-grid" id="summaryCardsGrid">
        <div class="card">
            <h3>إجمالي المصاريف العامة</h3>
            <div class="value" id="totalExpense">0 ريـال</div>
        </div>
    </div>

    <h3 style="margin-bottom: 12px; color: #0b5ed7; font-size: 16px;">🔄 تسوية الحسابات النهائية وملخصات الشركاء</h3>
    <div style="margin-bottom: 20px;" class="cards-grid" id="settlementCards"></div>
</div>

<!-- نافذة إنشاء مشروع جديد -->
<div class="modal-overlay" id="newAccountModal">
    <div class="modal-content">
        <button class="close-modal" type="button" onclick="closeNewAccountModal()">✕</button>
        <h2 style="color: #20c997; margin-bottom: 15px;">📁 إنشاء كشف حساب / مشروع جديد</h2>
        <div class="form-container" style="background: #e6fcf5;">
            <label style="font-weight: bold; font-size: 13px; display: block; margin-bottom: 5px;">اسم كشف الحساب أو المشروع الجديد:</label>
            <input type="text" id="newAccountName" placeholder="مثال: مشروع بئر إضافي / استصلاح أرض..." style="margin-bottom: 12px;">
            
            <label style="font-weight: bold; font-size: 13px; display: block; margin-bottom: 5px;">وصف مختصر:</label>
            <input type="text" id="newAccountDesc" placeholder="تفاصيل الشركاء والنشاط..." style="margin-bottom: 15px;">
        </div>
        <button type="button" onclick="createNewAccount()" style="background-color: #20c997; width: 100%; padding: 12px; font-size: 15px;">🚀 إنشاء والانتقال إليه فوراً</button>
    </div>
</div>

<!-- نافذة إعدادات الشركاء -->
<div class="modal-overlay" id="setupModal">
    <div class="modal-content">
        <button class="close-modal" type="button" onclick="closeSetupModal()">✕</button>
        <h2 style="color: #6f42c1; margin-bottom: 15px;">⚙️ إعدادات المشروع والتحكم بالشركاء</h2>
        
        <div class="form-container" style="background: #f3e8ff;">
            <label style="font-weight: bold; font-size: 13px; display: block; margin-bottom: 5px;">اسم المشروع أو النشاط:</label>
            <input type="text" id="setupProjectName" placeholder="أدخل اسم المشروع..." style="margin-bottom: 12px;">
            
            <label style="font-weight: bold; font-size: 13px; display: block; margin-bottom: 5px;">وصف المشروع:</label>
            <input type="text" id="setupProjectDesc" placeholder="تفاصيل الشركاء..." style="margin-bottom: 15px;">
        </div>

        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
            <h3 style="font-size: 14px; color: #333; margin: 0;">قائمة الشركاء والنسب:</h3>
            <span id="totalRateIndicator" style="font-size: 12px; font-weight: bold; color: #198754;">المجموع: 100%</span>
        </div>
        
        <div id="partnersInputsContainer"></div>
        <button type="button" onclick="addPartnerInputRow()" style="background-color: #0d6efd; margin-top: 5px; margin-bottom: 20px; padding: 8px;">➕ إضافة شريك جديد</button>

        <div style="display: flex; gap: 10px;">
            <button type="button" onclick="saveProjectSettings()" style="background-color: #6f42c1; flex: 1; padding: 12px;">💾 حفظ التعديلات</button>
            <button type="button" onclick="resetAllProjectData()" style="background-color: #dc3545; padding: 12px;">🔄 استعادة الافتراضي</button>
        </div>
    </div>
</div>

<!-- نافذة كشف الحساب العام المستقل -->
<div class="modal-overlay" id="generalModal">
    <div class="modal-content" id="generalModalContent">
        <button class="close-modal" type="button" onclick="closeGeneralModal()">✕</button>
        <h2 style="color: #0b5ed7; margin-bottom: 10px;">📋 نافذة كشف الحساب العام المستقل</h2>
        <div class="subtitle">إدارة وإنشاء كشفيات تفصيلية وعامة مستقلة تماماً</div>
        <div class="form-container">
            <h3>➕ إضافة بند جديد للكشف العام</h3>
            <input type="hidden" id="genEditIndex" value="-1">
            <div class="form-group" style="grid-template-columns: 1fr 2fr 1fr auto;">
                <input type="date" id="genDateInput">
                <input type="text" id="genDescInput" placeholder="تفاصيل البند العام...">
                <input type="number" id="genAmountInput" placeholder="المبلغ">
                <button type="button" onclick="saveGeneralItem()" style="background-color: #0b5ed7;">إضافة بند</button>
            </div>
        </div>
        <div class="table-responsive">
            <table>
                <thead><tr><th>التاريخ</th><th>التفاصيل</th><th>المبلغ</th><th>إجراءات</th></tr></thead>
                <tbody id="generalTableBody"></tbody>
            </table>
        </div>
        <div style="margin-top: 15px; background: #e7f1ff; padding: 12px; border-radius: 8px; display: flex; justify-content: space-between; align-items: center;">
            <span style="font-weight: bold; color: #0b5ed7;">إجمالي الكشف العام:</span>
            <span id="genTotalValue" style="font-weight: bold; font-size: 16px; color: #0b5ed7;">0 ريـال</span>
        </div>
        <div style="margin-top: 20px; display: flex; gap: 10px;">
            <button type="button" class="btn-pdf" onclick="exportGeneralPDF()" style="flex: 1; padding: 10px;">📄 تصدير الكشف العام PDF</button>
            <button type="button" class="btn-whatsapp" onclick="sendGeneralWhatsApp()" style="flex: 1; padding: 10px;">💬 إرسال الكشف العام واتساب</button>
        </div>
    </div>
</div>

<!-- نافذة التكاليف والخسائر -->
<div class="modal-overlay" id="lossesModal">
    <div class="modal-content" id="lossesModalContent">
        <button class="close-modal" type="button" onclick="closeLossesModal()">✕</button>
        <h2 style="color: #fd7e14; margin-bottom: 10px;">📉 سجل التكاليف والخسائر العامة</h2>
        <div class="subtitle">المصاريف النثرية أو الخسائر المعزولة</div>
        <div class="form-container">
            <h3>➕ إضافة بند تكلفة أو خسارة</h3>
            <input type="hidden" id="lossEditIndex" value="-1">
            <div class="form-group" style="grid-template-columns: 1fr 2fr 1fr auto;">
                <input type="date" id="lossDateInput">
                <input type="text" id="lossDescInput" placeholder="التفاصيل...">
                <input type="number" id="lossAmountInput" placeholder="المبلغ">
                <button type="button" onclick="saveLossItem()" style="background-color: #fd7e14;">إضافة</button>
            </div>
        </div>
        <div class="table-responsive">
            <table>
                <thead><tr><th>التاريخ</th><th>التفاصيل</th><th>المبلغ</th><th>إجراءات</th></tr></thead>
                <tbody id="lossesTableBody"></tbody>
            </table>
        </div>
        <div style="margin-top: 15px; background: #fff3cd; padding: 12px; border-radius: 8px; display: flex; justify-content: space-between; align-items: center;">
            <span style="font-weight: bold; color: #856404;">الإجمالي:</span>
            <span id="lossesTotalValue" style="font-weight: bold; font-size: 16px; color: #856404;">0 ريـال</span>
        </div>
        <div style="margin-top: 20px; display: flex; gap: 10px;">
            <button type="button" class="btn-pdf" onclick="exportLossesPDF()" style="flex: 1; padding: 10px;">📄 تصدير PDF</button>
            <button type="button" class="btn-whatsapp" onclick="sendLossesWhatsApp()" style="flex: 1; padding: 10px;">💬 واتساب</button>
        </div>
    </div>
</div>

<script>
    const defaultWellPartners = [
        { name: "أولاد عبده أحمد حسين", shareRate: 0.375 },
        { name: "أولاد صالح أحمد حسين", shareRate: 0.375 },
        { name: "أولاد حسين أحمد حسين", shareRate: 0.25 }
    ];

    const defaultWellExpenses = [
        { date: "2026-08-13", desc: "بترول معبر العشي", amount: 5000, payer: "صندوق البئر (مشترك)" },
        { date: "2026-08-14", desc: "بترول لاعسم", amount: 4000, payer: "صندوق البئر (مشترك)" },
        { date: "2026-08-15", desc: "اجور الونش توصيل المواصير", amount: 5000, payer: "صندوق البئر (مشترك)" },
        { date: "2026-08-16", desc: "محمد عامر قص المحافضه", amount: 15000, payer: "صندوق البئر (مشترك)" },
        { date: "2026-08-16", desc: "مع انور علي عبده يشل مسامير", amount: 4000, payer: "صندوق البئر (مشترك)" },
        { date: "2026-08-16", desc: "شلك 3", amount: 4500, payer: "صندوق البئر (مشترك)" },
        { date: "2026-08-16", desc: "مسامير جديد 2.5", amount: 30000, payer: "صندوق البئر (مشترك)" },
        { date: "2026-08-17", desc: "مع حواله مع الونش", amount: 50000, payer: "صندوق البئر (مشترك)" },
        { date: "2026-08-19", desc: "مع الصانع حق وزن الراس", amount: 10000, payer: "صندوق البئر (مشترك)" },
        { date: "2026-08-20", desc: "من حسين عبده يوم المروحه في معبر", amount: 5000, payer: "أولاد حسين أحمد حسين" },
        { date: "2026-08-20", desc: "بترول يوم دخلو المروحه معبر", amount: 5000, payer: "صندوق البئر (مشترك)" },
        { date: "2026-08-20", desc: "عمال نقل وتنزيل المواصير", amount: 13000, payer: "صندوق البئر (مشترك)" },
        { date: "2026-08-20", desc: "من محمد صالح بيد محمد حسين - مشروبات وصابون", amount: 4500, payer: "أولاد صالح أحمد حسين" },
        { date: "2026-08-20", desc: "من محمد صالح جبن وطحينيه وخبز", amount: 6000, payer: "أولاد صالح أحمد حسين" },
        { date: "2026-08-20", desc: "من محمد صالح صنافير", amount: 1000, payer: "أولاد صالح أحمد حسين" },
        { date: "2026-08-20", desc: "بترول سياره محمد صالح ادا المواصير من بير منصور", amount: 3000, payer: "أولاد صالح أحمد حسين" },
        { date: "2026-08-20", desc: "بترول لسياره حسين عبده حق توصيل السيبه", amount: 3000, payer: "أولاد حسين أحمد حسين" },
        { date: "2026-08-20", desc: "مجارحه يد عبدالله صالح", amount: 10000, payer: "صندوق البئر (مشترك)" },
        { date: "2026-08-20", desc: "قات من حسين عبده للونش", amount: 12500, payer: "أولاد حسين أحمد حسين" },
        { date: "2026-08-20", desc: "ريشه قص المحافضه", amount: 1000, payer: "صندوق البئر (مشترك)" },
        { date: "2026-08-20", desc: "من حسين عبده صاحب عسم", amount: 2000, payer: "أولاد حسين أحمد حسين" },
        { date: "2026-08-20", desc: "بترول عبدالله حسين مشوارين معبر", amount: 10000, payer: "صندوق البئر (مشترك)" },
        { date: "2026-08-20", desc: "مجارحه عبدالله صالح", amount: 6000, payer: "صندوق البئر (مشترك)" },
        { date: "2026-08-20", desc: "بترول تبع البير سيره جيه", amount: 5000, payer: "صندوق البئر (مشترك)" },
        { date: "2026-08-21", desc: "بترول حسين عبده للمزرعه يوم اسعفو عبدالله صالح", amount: 2000, payer: "أولاد حسين أحمد حسين" },
        { date: "2026-08-24", desc: "اجره معي انا واحمد علي عبده حق لبنه واخراج منافس البير", amount: 15000, payer: "صندوق البئر (مشترك)" },
        { date: "2026-08-26", desc: "قات من حسين عبده للونش اخر يوم", amount: 7500, payer: "أولاد حسين أحمد حسين" },
        { date: "2026-08-26", desc: "عشا من احمد عبده للونش", amount: 3000, payer: "أولاد عبده أحمد حسين" },
        { date: "2026-09-03", desc: "بترول تبع موضوع المناقله لاعند صالح عبد الله", amount: 5000, payer: "صندوق البئر (مشترك)" }
    ];

    let allProjects = JSON.parse(localStorage.getItem('all_accounting_projects_v2')) || {
        "well_zilhwa": {
            name: "حسابات مخاريج البئر محلي",
            desc: "تخص البئر المشتركة بين: أولاد عبده أحمد حسين، أولاد صالح أحمد حسين، وأولاد حسين أحمد حسين",
            partners: defaultWellPartners,
            expenses: defaultWellExpenses,
            partnerPayments: [],
            lossesItems: [],
            generalStatement: [{ date: "2026-09-01", desc: "بند افتتاح الكشف العام", amount: 10000 }]
        }
    };

    let currentProjectId = localStorage.getItem('current_accounting_project_id_v2') || "well_zilhwa";
    if (!allProjects[currentProjectId]) { currentProjectId = Object.keys(allProjects)[0]; }

    let projectConfig = allProjects[currentProjectId];
    let expenses = projectConfig.expenses || [];
    let partnerPayments = projectConfig.partnerPayments || [];
    let lossesItems = projectConfig.lossesItems || [];
    let generalStatementItems = projectConfig.generalStatement || [];

    let globalTotal = 0;
    let partnerSharesGlobal = {};
    let partnerPaidGlobal = {};
    let partnerDirectPaymentsGlobal = {};
    let lossesTotalGlobal = 0;
    let generalTotalGlobal = 0;

    document.getElementById('dateInput').valueAsDate = new Date();
    document.getElementById('payDateInput').valueAsDate = new Date();
    document.getElementById('lossDateInput').valueAsDate = new Date();
    document.getElementById('genDateInput').valueAsDate = new Date();

    function saveData() {
        projectConfig.expenses = expenses;
        projectConfig.partnerPayments = partnerPayments;
        projectConfig.lossesItems = lossesItems;
        projectConfig.generalStatement = generalStatementItems;
        allProjects[currentProjectId] = projectConfig;
        localStorage.setItem('all_accounting_projects_v2', JSON.stringify(allProjects));
        localStorage.setItem('current_accounting_project_id_v2', currentProjectId);
    }

    function updateProjectSelector() {
        const selector = document.getElementById('projectSelector');
        selector.innerHTML = '';
        for (let id in allProjects) {
            let opt = document.createElement('option');
            opt.value = id;
            opt.innerText = allProjects[id].name;
            if (id === currentProjectId) opt.selected = true;
            selector.appendChild(opt);
        }
    }

    function switchProject(id) {
        if (allProjects[id]) {
            currentProjectId = id;
            projectConfig = allProjects[id];
            expenses = projectConfig.expenses || [];
            partnerPayments = projectConfig.partnerPayments || [];
            lossesItems = projectConfig.lossesItems || [];
            generalStatementItems = projectConfig.generalStatement || [];
            localStorage.setItem('current_accounting_project_id_v2', currentProjectId);
            renderSystem();
        }
    }

    function openNewAccountModal() {
        document.getElementById('newAccountName').value = '';
        document.getElementById('newAccountDesc').value = '';
        document.getElementById('newAccountModal').style.display = 'flex';
    }
    function closeNewAccountModal() { document.getElementById('newAccountModal').style.display = 'none'; }

    function createNewAccount() {
        const name = document.getElementById('newAccountName').value.trim();
        const desc = document.getElementById('newAccountDesc').value.trim();
        if (!name) { alert('الرجاء إدخال اسم المشروع.'); return; }

        let newId = 'proj_' + Date.now();
        allProjects[newId] = {
            name: name,
            desc: desc || "كشف حساب جديد للمصاريف والتكاليف",
            partners: [
                { name: "الشريك الأول", shareRate: 0.50 },
                { name: "الشريك الثاني", shareRate: 0.50 }
            ],
            expenses: [],
            partnerPayments: [],
            lossesItems: [],
            generalStatement: []
        };

        currentProjectId = newId;
        projectConfig = allProjects[currentProjectId];
        expenses = projectConfig.expenses;
        partnerPayments = projectConfig.partnerPayments;
        lossesItems = projectConfig.lossesItems;
        generalStatementItems = projectConfig.generalStatement;

        saveData();
        closeNewAccountModal();
        renderSystem();
        alert(`تم إنشاء المشروع "${name}" والانتقال إليه بنجاح!`);
    }

    function renderSystem() {
        updateProjectSelector();
        document.getElementById('displayProjectName').innerText = projectConfig.name;
        document.getElementById('displayProjectDesc').innerText = projectConfig.desc;
        updateDropdowns();

        const tbody = document.getElementById('expenseTableBody');
        tbody.innerHTML = '';
        let totalExpense = 0;
        let partnerPaidFromExpenses = {};
        projectConfig.partners.forEach(p => partnerPaidFromExpenses[p.name] = 0);

        expenses.forEach((item, index) => {
            totalExpense += item.amount;
            if (partnerPaidFromExpenses.hasOwnProperty(item.payer)) {
                partnerPaidFromExpenses[item.payer] += item.amount;
            }

            const row = document.createElement('tr');
            let payerBadgeClass = item.payer.includes('مشترك') ? 'badge' : 'badge badge-partner';
            row.innerHTML = `
                <td style="white-space: nowrap;">${item.date}</td>
                <td>${item.desc}</td>
                <td style="white-space: nowrap;"><strong>${item.amount.toLocaleString()} ريـال</strong></td>
                <td><span class="${payerBadgeClass}">${item.payer}</span></td>
                <td>
                    <div class="action-buttons-cell">
                        <button type="button" class="btn-action edit-btn" onclick="editExpense(${index})">تعديل</button>
                        <button type="button" class="btn-action delete-btn" onclick="deleteExpense(${index})">حذف</button>
                    </div>
                </td>
            `;
            tbody.appendChild(row);
        });

        globalTotal = totalExpense;
        document.getElementById('totalExpense').innerText = totalExpense.toLocaleString() + " ريـال";

        const summaryGrid = document.getElementById('summaryCardsGrid');
        summaryGrid.innerHTML = `
            <div class="card">
                <h3>إجمالي المصاريف العامة</h3>
                <div class="value" id="totalExpense">${totalExpense.toLocaleString()} ريـال</div>
            </div>
        `;

        partnerSharesGlobal = {};
        projectConfig.partners.forEach(p => {
            let shareAmount = totalExpense * p.shareRate;
            partnerSharesGlobal[p.name] = shareAmount;
            
            const card = document.createElement('div');
            card.className = 'card partner-card';
            card.innerHTML = `
                <h3>حصة ${p.name} (${(p.shareRate * 100)}%)</h3>
                <div class="value">${shareAmount.toLocaleString(undefined, {maximumFractionDigits: 2})} ريـال</div>
            `;
            summaryGrid.appendChild(card);
        });

        partnerPaidGlobal = partnerPaidFromExpenses;

        const payTbody = document.getElementById('paymentTableBody');
        payTbody.innerHTML = '';
        let partnerDirectPayments = {};
        projectConfig.partners.forEach(p => partnerDirectPayments[p.name] = 0);

        partnerPayments.forEach((pItem, pIndex) => {
            if (partnerDirectPayments.hasOwnProperty(pItem.partner)) {
                partnerDirectPayments[pItem.partner] += pItem.amount;
            }

            const row = document.createElement('tr');
            row.innerHTML = `
                <td style="white-space: nowrap;">${pItem.date}</td>
                <td>${pItem.desc}</td>
                <td style="white-space: nowrap;"><strong style="color: #0d6efd;">${pItem.amount.toLocaleString()} ريـال</strong></td>
                <td><span class="badge badge-payment">${pItem.partner}</span></td>
                <td>
                    <div class="action-buttons-cell">
                        <button type="button" class="btn-action edit-btn" onclick="editPartnerPayment(${pIndex})">تعديل</button>
                        <button type="button" class="btn-action delete-btn" onclick="deletePartnerPayment(${pIndex})">حذف</button>
                    </div>
                </td>
            `;
            payTbody.appendChild(row);
        });

        partnerDirectPaymentsGlobal = partnerDirectPayments;

        const settlementContainer = document.getElementById('settlementCards');
        settlementContainer.innerHTML = '';

        projectConfig.partners.forEach(p => {
            let requiredShare = partnerSharesGlobal[p.name];
            let paidFromExp = partnerPaidFromExpenses[p.name] || 0;
            let directPaid = partnerDirectPayments[p.name] || 0;
            let totalPaidByPartner = paidFromExp + directPaid;
            let net = requiredShare - totalPaidByPartner; 

            let statusText = "";
            let statusColor = "";
            if (net > 0) {
                statusText = `المطلوب دفعه عليه: ${net.toLocaleString(undefined, {maximumFractionDigits: 2})} ريـال`;
                statusColor = "#0d6efd";
            } else if (net < 0) {
                statusText = `له في ذمة المشروع: ${Math.abs(net).toLocaleString(undefined, {maximumFractionDigits: 2})} ريـال`;
                statusColor = "#198754";
            } else {
                statusText = `حسابه خالص تماماً`;
                statusColor = "#6c757d";
            }

            const card = document.createElement('div');
            card.className = 'settlement-card';
            card.style.borderRightColor = statusColor;
            card.innerHTML = `
                <h3 style="color: #0b5ed7; margin-bottom: 5px; font-size: 14px;">👤 ${p.name}</h3>
                <div style="font-size: 12px; color: #555; margin-bottom: 2px;">حسبته المقررة: <strong>${requiredShare.toLocaleString(undefined, {maximumFractionDigits: 2})} ريـال</strong></div>
                <div style="font-size: 12px; color: #555; margin-bottom: 2px;">ما دفعه: <strong>${totalPaidByPartner.toLocaleString()} ريـال</strong></div>
                <div class="value" style="color: ${statusColor}; font-size: 13px; margin-bottom: 8px;">${statusText}</div>
                <button type="button" class="btn-share-whatsapp" onclick="sendPartnerSummary('${p.name}')">📤 إرسال الملخص عبر واتساب</button>
            `;
            settlementContainer.appendChild(card);
        });

        renderLossesTable();
        renderGeneralTable();
        saveData();
    }

    function updateDropdowns() {
        let defaultBoxName = projectConfig.name.includes('بئر') ? 'صندوق البئر (مشترك)' : 'صندوق المشروع (مشترك)';
        const payerSelect = document.getElementById('payerSelect');
        payerSelect.innerHTML = `<option value="${defaultBoxName}">${defaultBoxName}</option>`;
        const payPartnerSelect = document.getElementById('payPartnerSelect');
        payPartnerSelect.innerHTML = '';

        projectConfig.partners.forEach(p => {
            payerSelect.innerHTML += `<option value="${p.name}">${p.name}</option>`;
            payPartnerSelect.innerHTML += `<option value="${p.name}">${p.name}</option>`;
        });
    }

    // استيراد ملفات Excel
    function importExcelData(event) {
        const file = event.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const jsonRows = XLSX.utils.sheet_to_json(worksheet, {header: 1});

                let importedCount = 0;
                let defaultBoxName = projectConfig.name.includes('بئر') ? 'صندوق البئر (مشترك)' : 'صندوق المشروع (مشترك)';
                for (let i = 1; i < jsonRows.length; i++) {
                    let row = jsonRows[i];
                    if (row && row.length >= 3) {
                        let dateVal = row[0] ? String(row[0]) : new Date().toISOString().split('T')[0];
                        let descVal = row[1] ? String(row[1]) : 'بند مستورد';
                        let amountVal = parseFloat(row[2]) || 0;
                        let payerVal = row[3] ? String(row[3]) : defaultBoxName;

                        if (amountVal > 0) {
                            expenses.push({ date: dateVal, desc: descVal, amount: amountVal, payer: payerVal });
                            importedCount++;
                        }
                    }
                }
                if (importedCount > 0) { alert(`تم استيراد ${importedCount} بند بنجاح من إكسل!`); renderSystem(); }
                else { alert('لم يتم العثور على بيانات صالحة.'); }
            } catch (err) { alert('خطأ في قراءة ملف الإكسل.'); }
            event.target.value = '';
        };
        reader.readAsArrayBuffer(file);
    }

    // استيراد ملفات Word
    async function importWordData(event) {
        const file = event.target.files[0];
        if (!file) return;
        try {
            const arrayBuffer = await file.arrayBuffer();
            const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
            const lines = result.value.split('\n');
            let importedCount = 0;
            const todayStr = new Date().toISOString().split('T')[0];
            let defaultBoxName = projectConfig.name.includes('بئر') ? 'صندوق البئر (مشترك)' : 'صندوق المشروع (مشترك)';

            lines.forEach(line => {
                let trimmed = line.trim();
                if(trimmed.length > 3) {
                    let numbers = trimmed.match(/[\d,.]+/g);
                    if(numbers && numbers.length > 0) {
                        let amount = parseFloat(numbers[numbers.length - 1].replace(/,/g, ''));
                        if(!isNaN(amount) && amount > 10) {
                            let desc = trimmed.replace(numbers[numbers.length - 1], '').trim();
                            if(desc.length < 2) desc = "بند مستورد من Word";
                            expenses.push({ date: todayStr, desc: desc, amount: amount, payer: defaultBoxName });
                            importedCount++;
                        }
                    }
                }
            });

            if (importedCount > 0) { alert(`تم استيراد ${importedCount} بند من Word بنجاح!`); renderSystem(); }
            else { alert('الملف لا يحتوي على أسطر مبالغ واضحة.'); }
        } catch (err) { alert('حدث خطأ أثناء قراءة ملف الـ Word.'); }
        event.target.value = '';
    }

    // استيراد ملفات PDF
    async function importPdfData(event) {
        const file = event.target.files[0];
        if (!file) return;
        try {
            const arrayBuffer = await file.arrayBuffer();
            const pdfDoc = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
            let fullText = "";
            for (let i = 1; i <= pdfDoc.numPages; i++) {
                const page = await pdfDoc.getPage(i);
                const textContent = await page.getTextContent();
                fullText += textContent.items.map(item => item.str).join(" ") + "\n";
            }
            const lines = fullText.split('\n');
            let importedCount = 0;
            const todayStr = new Date().toISOString().split('T')[0];
            let defaultBoxName = projectConfig.name.includes('بئر') ? 'صندوق البئر (مشترك)' : 'صندوق المشروع (مشترك)';

            lines.forEach(line => {
                let trimmed = line.trim();
                if(trimmed.length > 1) {
                    let numbers = trimmed.match(/[\d,.]+/g);
                    if(numbers && numbers.length > 0) {
                        let amount = parseFloat(numbers[numbers.length - 1].replace(/,/g, ''));
                        if(!isNaN(amount) && amount > 0) {
                            let desc = trimmed.replace(numbers[numbers.length - 1], '').trim();
                            if(desc.length < 2) desc = "مستند PDF: " + file.name;
                            expenses.push({ date: todayStr, desc: desc, amount: amount, payer: defaultBoxName });
                            importedCount++;
                        }
                    }
                }
            });

            if (importedCount > 0) { alert(`تم قراءة واستخراج ${importedCount} بند من PDF بنجاح!`); renderSystem(); }
            else { alert('تعذر قراءة نصوص واضحة من ملف الـ PDF.'); }
        } catch (err) { alert('حدث خطأ أثناء تحليل ملف الـ PDF.'); }
        event.target.value = '';
    }

    function openSetupModal() {
        document.getElementById('setupProjectName').value = projectConfig.name;
        document.getElementById('setupProjectDesc').value = projectConfig.desc;
        const container = document.getElementById('partnersInputsContainer');
        container.innerHTML = '';
        projectConfig.partners.forEach(p => { appendPartnerInputRow(p.name, p.shareRate * 100); });
        document.getElementById('setupModal').style.display = 'flex';
    }
    function closeSetupModal() { document.getElementById('setupModal').style.display = 'none'; }

    function appendPartnerInputRow(name = '', rate = '') {
        const container = document.getElementById('partnersInputsContainer');
        const row = document.createElement('div');
        row.className = 'partner-input-row';
        row.innerHTML = `
            <input type="text" class="p-name-input" placeholder="اسم الشريك" value="${name}" style="flex: 2;">
            <input type="number" class="p-rate-input" placeholder="النسبة %" value="${rate}" style="flex: 1;" oninput="calculateTotalRatePreview()">
            <button type="button" onclick="this.parentElement.remove(); calculateTotalRatePreview();" style="background-color: #dc3545; width: 40px; padding: 10px;">✕</button>
        `;
        container.appendChild(row);
        calculateTotalRatePreview();
    }
    function addPartnerInputRow() { appendPartnerInputRow('', ''); }

    function calculateTotalRatePreview() {
        const rateInputs = document.querySelectorAll('.p-rate-input');
        let total = 0;
        rateInputs.forEach(input => { total += parseFloat(input.value) || 0; });
        const indicator = document.getElementById('totalRateIndicator');
        indicator.innerText = `المجموع الكلي: ${total}%`;
        indicator.style.color = Math.abs(total - 100) < 0.1 ? '#198754' : '#dc3545';
    }

    function saveProjectSettings() {
        const name = document.getElementById('setupProjectName').value.trim();
        const desc = document.getElementById('setupProjectDesc').value.trim();
        const nameInputs = document.querySelectorAll('.p-name-input');
        const rateInputs = document.querySelectorAll('.p-rate-input');

        let newPartners = [];
        let totalRate = 0;
        for (let i = 0; i < nameInputs.length; i++) {
            let pName = nameInputs[i].value.trim();
            let pRate = parseFloat(rateInputs[i].value);
            if (pName && !isNaN(pRate)) {
                newPartners.push({ name: pName, shareRate: pRate / 100 });
                totalRate += pRate;
            }
        }
        if (!name || newPartners.length === 0) { alert('الرجاء إدخال اسم المشروع وشريك واحد على الأقل مع نسبته.'); return; }
        if (Math.abs(totalRate - 100) > 0.1) { alert(`مجموع نسب الشركاء الحالي هو (${totalRate}%). يجب أن يكون المجموع 100% تماماً.`); return; }

        projectConfig.name = name;
        projectConfig.desc = desc;
        projectConfig.partners = newPartners;
        closeSetupModal();
        renderSystem();
    }

    function resetAllProjectData() {
        if(confirm('هل تريد استعادة البيانات الافتراضية الأصلية؟')) {
            projectConfig.partners = defaultWellPartners;
            projectConfig.expenses = defaultWellExpenses;
            projectConfig.partnerPayments = [];
            projectConfig.lossesItems = [];
            projectConfig.generalStatement = [{ date: "2026-09-01", desc: "بند افتتاح الكشف العام", amount: 10000 }];
            expenses = projectConfig.expenses;
            partnerPayments = projectConfig.partnerPayments;
            lossesItems = projectConfig.lossesItems;
            generalStatementItems = projectConfig.generalStatement;
            closeSetupModal();
            renderSystem();
        }
    }

    // إدارة المصاريف الرئيسية
    function saveExpense() {
        const editIndex = parseInt(document.getElementById('editIndex').value);
        const date = document.getElementById('dateInput').value;
        const desc = document.getElementById('descInput').value;
        const amount = parseFloat(document.getElementById('amountInput').value);
        const payer = document.getElementById('payerSelect').value;

        if (!date || !desc || isNaN(amount) || amount <= 0) { alert('الرجاء إدخال بيانات صحيحة'); return; }

        if (editIndex === -1) { expenses.push({ date, desc, amount, payer }); }
        else {
            expenses[editIndex] = { date, desc, amount, payer };
            document.getElementById('editIndex').value = -1;
            document.getElementById('formTitle').innerText = '➕ إضافة مصروف أو تكلفة جديدة';
            document.getElementById('saveBtn').innerText = 'إضافة القيد';
        }
        document.getElementById('descInput').value = '';
        document.getElementById('amountInput').value = '';
        document.getElementById('dateInput').valueAsDate = new Date();
        renderSystem();
    }

    function editExpense(index) {
        const item = expenses[index];
        document.getElementById('dateInput').value = item.date;
        document.getElementById('descInput').value = item.desc;
        document.getElementById('amountInput').value = item.amount;
        document.getElementById('payerSelect').value = item.payer;
        document.getElementById('editIndex').value = index;
        document.getElementById('formTitle').innerText = '✏️ تعديل قيد المصروف';
        document.getElementById('saveBtn').innerText = 'حفظ التعديل';
        window.scrollTo({ top: document.getElementById('formContainer').offsetTop, behavior: 'smooth' });
    }

    function deleteExpense(index) {
        if(confirm('حذف هذا السجل؟')) { expenses.splice(index, 1); renderSystem(); }
    }

    // إدارة الدفعات النقدية
    function savePartnerPayment() {
        const editIndex = parseInt(document.getElementById('editPaymentIndex').value);
        const date = document.getElementById('payDateInput').value;
        const desc = document.getElementById('payDescInput').value;
        const amount = parseFloat(document.getElementById('payAmountInput').value);
        const partner = document.getElementById('payPartnerSelect').value;

        if (!date || !desc || isNaN(amount) || amount <= 0) { alert('الرجاء إدخال بيانات الدفعة بدقة'); return; }

        if (editIndex === -1) { partnerPayments.push({ date, desc, amount, partner }); }
        else {
            partnerPayments[editIndex] = { date, desc, amount, partner };
            document.getElementById('editPaymentIndex').value = -1;
            document.getElementById('paymentFormTitle').innerText = '💵 تسجيل دفعة نقدية / تسديد من شريك';
            document.getElementById('savePaymentBtn').innerText = 'تسجيل الدفعة';
        }
        document.getElementById('payDescInput').value = '';
        document.getElementById('payAmountInput').value = '';
        document.getElementById('payDateInput').valueAsDate = new Date();
        renderSystem();
    }

    function editPartnerPayment(index) {
        const item = partnerPayments[index];
        document.getElementById('payDateInput').value = item.date;
        document.getElementById('payDescInput').value = item.desc;
        document.getElementById('payAmountInput').value = item.amount;
        document.getElementById('payPartnerSelect').value = item.partner;
        document.getElementById('editPaymentIndex').value = index;
        document.getElementById('paymentFormTitle').innerText = '✏️ تعديل الدفعة النقدية';
        document.getElementById('savePaymentBtn').innerText = 'حفظ التعديل';
        window.scrollTo({ top: document.getElementById('paymentFormContainer').offsetTop, behavior: 'smooth' });
    }

    function deletePartnerPayment(index) {
        if(confirm('حذف هذه الدفعة؟')) { partnerPayments.splice(index, 1); renderSystem(); }
    }

    // نافذة الكشف العام المستقل
    function openGeneralModal() { document.getElementById('generalModal').style.display = 'flex'; renderGeneralTable(); }
    function closeGeneralModal() { document.getElementById('generalModal').style.display = 'none'; }

    function renderGeneralTable() {
        const tbody = document.getElementById('generalTableBody');
        tbody.innerHTML = '';
        let total = 0;
        generalStatementItems.forEach((item, index) => {
            total += item.amount;
            tbody.innerHTML += `
                <tr>
                    <td>${item.date}</td>
                    <td>${item.desc}</td>
                    <td><strong>${item.amount.toLocaleString()} ريـال</strong></td>
                    <td>
                        <button type="button" class="btn-action edit-btn" onclick="editGeneralItem(${index})">تعديل</button>
                        <button type="button" class="btn-action delete-btn" onclick="deleteGeneralItem(${index})">حذف</button>
                    </td>
                </tr>`;
        });
        generalTotalGlobal = total;
        document.getElementById('genTotalValue').innerText = total.toLocaleString() + " ريـال";
        saveData();
    }

    function saveGeneralItem() {
        const editIndex = parseInt(document.getElementById('genEditIndex').value);
        const date = document.getElementById('genDateInput').value;
        const desc = document.getElementById('genDescInput').value;
        const amount = parseFloat(document.getElementById('genAmountInput').value);

        if (!date || !desc || isNaN(amount) || amount <= 0) return alert('الرجاء إدخال بيانات صحيحة للبند');
        if (editIndex === -1) generalStatementItems.push({ date, desc, amount });
        else { generalStatementItems[editIndex] = { date, desc, amount }; document.getElementById('genEditIndex').value = -1; }
        document.getElementById('genDescInput').value = '';
        document.getElementById('genAmountInput').value = '';
        renderGeneralTable();
    }

    function editGeneralItem(index) {
        const item = generalStatementItems[index];
        document.getElementById('genDateInput').value = item.date;
        document.getElementById('genDescInput').value = item.desc;
        document.getElementById('genAmountInput').value = item.amount;
        document.getElementById('genEditIndex').value = index;
    }

    function deleteGeneralItem(index) {
        if(confirm('حذف هذا البند؟')) { generalStatementItems.splice(index, 1); renderGeneralTable(); }
    }

    function exportGeneralPDF() { html2pdf().from(document.getElementById('generalModalContent')).save('كشف_الحساب_العام.pdf'); }

    function sendGeneralWhatsApp() {
        let msg = `📋 *كشف الحساب العام المستقل*\n`;
        generalStatementItems.forEach((item, i) => msg += `${i+1}. ${item.desc} (*${item.amount.toLocaleString()} ريـال*)\n`);
        msg += `💰 الإجمالي: *${generalTotalGlobal.toLocaleString()} ريـال*`;
        window.open(`https://api.whatsapp.com/send?text=${encodeURI(msg)}`, '_blank');
    }

    // نافذة الخسائر
    function openLossesModal() { document.getElementById('lossesModal').style.display = 'flex'; renderLossesTable(); }
    function closeLossesModal() { document.getElementById('lossesModal').style.display = 'none'; }

    function renderLossesTable() {
        const tbody = document.getElementById('lossesTableBody');
        tbody.innerHTML = '';
        let total = 0;
        lossesItems.forEach((item, index) => {
            total += item.amount;
            tbody.innerHTML += `
                <tr>
                    <td>${item.date}</td>
                    <td>${item.desc}</td>
                    <td><strong>${item.amount.toLocaleString()} ريـال</strong></td>
                    <td>
                        <button type="button" class="btn-action edit-btn" onclick="editLossItem(${index})">تعديل</button>
                        <button type="button" class="btn-action delete-btn" onclick="deleteLossItem(${index})">حذف</button>
                    </td>
                </tr>`;
        });
        lossesTotalGlobal = total;
        document.getElementById('lossesTotalValue').innerText = total.toLocaleString() + " ريـال";
        saveData();
    }

    function saveLossItem() {
        const editIndex = parseInt(document.getElementById('lossEditIndex').value);
        const date = document.getElementById('lossDateInput').value;
        const desc = document.getElementById('lossDescInput').value;
        const amount = parseFloat(document.getElementById('lossAmountInput').value);

        if (!date || !desc || isNaN(amount) || amount <= 0) return alert('خطأ في البيانات المدخلة');
        if (editIndex === -1) lossesItems.push({ date, desc, amount });
        else { lossesItems[editIndex] = { date, desc, amount }; document.getElementById('lossEditIndex').value = -1; }
        document.getElementById('lossDescInput').value = '';
        document.getElementById('lossAmountInput').value = '';
        renderLossesTable();
    }

    function editLossItem(index) {
        const item = lossesItems[index];
        document.getElementById('lossDateInput').value = item.date;
        document.getElementById('lossDescInput').value = item.desc;
        document.getElementById('lossAmountInput').value = item.amount;
        document.getElementById('lossEditIndex').value = index;
    }

    function deleteLossItem(index) {
        if(confirm('حذف هذا البند؟')) { lossesItems.splice(index, 1); renderLossesTable(); }
    }

    function exportLossesPDF() { html2pdf().from(document.getElementById('lossesModalContent')).save('سجل_الخسائر.pdf'); }

    function sendLossesWhatsApp() {
        let msg = `📉 *سجل التكاليف والخسائر العامة - ${projectConfig.name}*\n`;
        lossesItems.forEach((item, i) => msg += `${i+1}. ${item.desc} (*${item.amount.toLocaleString()} ريـال*)\n`);
        msg += `💰 الإجمالي: *${lossesTotalGlobal.toLocaleString()} ريـال*`;
        window.open(`https://api.whatsapp.com/send?text=${encodeURI(msg)}`, '_blank');
    }

    // إرسال ملخص الشريك عبر واتساب
    function sendPartnerSummary(partnerName) {
        let requiredShare = partnerSharesGlobal[partnerName];
        let paidExp = partnerPaidGlobal[partnerName] || 0;
        let directPaid = partnerDirectPaymentsGlobal[partnerName] || 0;
        let totalPaid = paidExp + directPaid;
        let net = requiredShare - totalPaid;

        let msg = `📊 *ملخص حساب المشروع: ${projectConfig.name}*\n`;
        msg += `👤 *الشريك:* ${partnerName}\n`;
        msg += `═══════════════════\n`;
        msg += `💰 إجمالي المصاريف: *${globalTotal.toLocaleString()} ريـال*\n`;
        msg += `📌 حصتك المقررة: *${requiredShare.toLocaleString(undefined, {maximumFractionDigits: 2})} ريـال*\n`;
        msg += `💵 ما دفعته: *${totalPaid.toLocaleString()} ريـال*\n`;
        
        if (net > 0) msg += `🔴 *المطلوب دفعه عليك:* *${net.toLocaleString(undefined, {maximumFractionDigits: 2})} ريـال*\n`;
        else if (net < 0) msg += `🟢 *لك في ذمة المشروع:* *${Math.abs(net).toLocaleString(undefined, {maximumFractionDigits: 2})} ريـال*\n`;
        else msg += `⚪ *حسابك خالص تماماً.*\n`;

        window.open(`https://api.whatsapp.com/send?text=${encodeURI(msg)}`, '_blank');
    }

    // تصدير PDF الرئيسي
    function exportPDF() {
        const hideList = [
            document.getElementById('actionBar'),
            document.getElementById('formContainer'),
            document.getElementById('paymentFormContainer'),
            document.getElementById('actionHeader'),
            document.getElementById('paymentActionHeader')
        ];
        hideList.forEach(el => el.classList.add('pdf-hidden'));
        document.querySelectorAll('.btn-action, .btn-share-whatsapp').forEach(el => el.classList.add('pdf-hidden'));

        html2pdf().from(document.getElementById('pdfContent')).set({
            margin: 10, filename: `${projectConfig.name}_كشف_حساب.pdf`,
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { scale: 2, useCORS: true },
            jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
        }).save().then(() => {
            hideList.forEach(el => el.classList.remove('pdf-hidden'));
            document.querySelectorAll('.btn-action, .btn-share-whatsapp').forEach(el => el.classList.remove('pdf-hidden'));
        });
    }

    function downloadPdfAndOpenWhatsApp() {
        exportPDF();
        setTimeout(() => {
            let msg = `📊 *كشف حساب مشروع: ${projectConfig.name}*\n💰 إجمالي المصاريف العامة: *${globalTotal.toLocaleString()} ريـال*\n\n_تم تنزيل ملف PDF بنجاح._`;
            window.open(`https://api.whatsapp.com/send?text=${encodeURI(msg)}`, '_blank');
        }, 1000);
    }

    renderSystem();
</script>

</body>
</html>
