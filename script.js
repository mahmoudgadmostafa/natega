let studentsData = [];
const EXCEL_URL = "https://docs.google.com/spreadsheets/d/1UWS_ZnKEzufnabYi6PT4-04gJUIZj7gL/export?format=xlsx";

async function loadExcelFromDrive() {
    try {
        let response = await fetch(EXCEL_URL);
        let data = await response.arrayBuffer();
        let workbook = XLSX.read(new Uint8Array(data), { type: "array" });

        studentsData = [];

        workbook.SheetNames.forEach(sheetName => {
            let sheet = workbook.Sheets[sheetName];
            let sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            if (sheetData.length > 1) {
                let headers = sheetData[0];
                let formattedData = sheetData.slice(1).map(row => {
                    let obj = {};
                    headers.forEach((header, i) => obj[header] = row[i]);
                    return obj;
                });

                studentsData.push({
                    sheetName: sheetName,
                    headers: headers,
                    data: formattedData,
                });
            }
        });

        console.log("📊 تم تحميل جميع البيانات بنجاح:", studentsData);
    } catch (error) {
        console.error("❌ خطأ أثناء تحميل البيانات.", error);
    }
}

function fetchStudentResult() {
    let studentId = document.getElementById("studentId").value.trim();
    if (!studentId) {
        alert("⚠️ يرجى إدخال رقم الطالب");
        return;
    }

    let foundResults = [];

    studentsData.forEach(sheet => {
        let student = sheet.data.find(s => String(s["رقم الطالب"]).trim() === studentId);
        if (student) {
            foundResults.push({
                sheetName: sheet.sheetName,
                headers: sheet.headers,
                student: student
            });
        }
    });

    let resultDiv = document.getElementById("result");
    resultDiv.innerHTML = "";

    if (foundResults.length > 0) {

        // ⭐ بيانات الطالب أعلى الجدول
        let studentInfo = foundResults[0].student;
        resultDiv.innerHTML += `
            <div class="student-info-box">
                <p><strong>📌 اسم الطالب:</strong> ${studentInfo["الاسم"] || "غير متوفر"}</p>
                <p><strong>📌 رقم الطالب:</strong> ${studentId}</p>
            </div>
        `;

        foundResults.forEach(result => {
            resultDiv.innerHTML += `
                <h3 class="sheet-title">📘 نتيجة الطالب - (${result.sheetName})</h3>
            `;

            let tableHTML = `
                <table class="result-table">
                    <thead>
                        <tr>
                            <th>المادة</th>
                            <th>الدرجة</th>
                        </tr>
                    </thead>
                    <tbody>
            `;

            result.headers.forEach(header => {
                if (
                    header !== "رقم الطالب" &&
                    header !== "الاسم" &&
                    result.student[header] !== undefined
                ) {
                    tableHTML += `
                        <tr>
                            <td>${header}</td>
                            <td>${result.student[header]}</td>
                        </tr>
                    `;
                }
            });

            tableHTML += `
                    </tbody>
                </table>
                <br>
            `;

            resultDiv.innerHTML += tableHTML;
        });

        // إظهار زر إدخال رقم جديد
        document.getElementById("resetBtn").style.display = "block";

    } else {
        alert("❌ لم يتم العثور على الطالب.");
    }
}

// إعادة التهيئة
function resetSearch() {
    document.getElementById("studentId").value = "";
    document.getElementById("result").innerHTML = "";
    document.getElementById("resetBtn").style.display = "none";
}

// تفريغ النتائج عند مسح الرقم
document.getElementById("studentId").addEventListener("input", function () {
    if (this.value.trim() === "") {
        document.getElementById("result").innerHTML = "";
        document.getElementById("resetBtn").style.display = "none";
    }
});

// تحميل البيانات عند فتح الصفحة
loadExcelFromDrive();
