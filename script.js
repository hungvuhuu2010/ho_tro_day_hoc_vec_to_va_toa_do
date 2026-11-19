/**

 * ỨNG DỤNG HỖ TRỢ DẠY - HỌC TỌA ĐỘ VECTƠ (MOBILE APP ENGINE)

 * Chuyên gia Thiết kế Trắc nghiệm & WebApp Mobile

 */



// ==========================================

// 1. TRẠNG THÁI HỆ THỐNG & KẾT NỐI LOCALSTORAGE

// ==========================================

let isChallenge = false;

let selectedTopics = [1, 2, 3, 4, 5, 6];

let score = 0;

let streak = 0;

let highScore = parseInt(localStorage.getItem('vector_app_highscore')) || 0;

let currentQuestion = null;


let excelTheoryData = {};
// Khởi tạo app khi DOM sẵn sàng

document.addEventListener('DOMContentLoaded', () => {

    updateHighScoreDisplay();

});


// Gọi hàm này khi khởi động app
async function loadTheoryFromExcel() {
    try {
        let response = await fetch('thong_tin_cong_thuc.xlsx');
        let data = await response.arrayBuffer();
        let workbook = XLSX.read(data, {type: 'array'});
        
        // Lấy sheet đầu tiên
        let firstSheetName = workbook.SheetNames[0];
        let worksheet = workbook.Sheets[firstSheetName];
        
        // Chuyển sheet thành dạng JSON (mảng các object)
        let rows = XLSX.utils.sheet_to_json(worksheet);
        
        // Lập chỉ mục tra cứu theo topic_id và sub_type
        rows.forEach(row => {
            let key = `${row.topic_id}_${row.sub_type}`;
            excelTheoryData[key] = {
                title: row.ten_dang_toan,
                formula: row.cong_thuc_chuan,
                guide: row.huong_dan_giai,
                commonMistake: row.loi_sai_pho_bien
            };
        });
        console.log("Đã tải thành công dữ liệu lý thuyết từ Excel!");
    } catch (e) {
        console.error("Không thể đọc file Excel lý thuyết:", e);
    }
}

// Cập nhật hiển thị kỷ lục

function updateHighScoreDisplay() {

    const hsEl = document.getElementById('highScoreDisplay');

    if (hsEl) {

        hsEl.innerText = `Kỷ lục: ${highScore}`;

    }

}



// Rung phản hồi haptic cho thiết bị di động

function triggerHaptic(type) {

    if ('navigator' in window && 'vibrate' in navigator) {

        if (type === 'correct') {

            navigator.vibrate([40, 30, 40]); // Rung nhẹ 2 nhịp khi đúng

        } else if (type === 'wrong') {

            navigator.vibrate(200); // Rung dài khi sai

        } else if (type === 'win') {

            navigator.vibrate([100, 50, 100, 50, 200]); // Rung chúc mừng

        }

    }

}



// ==========================================

// 2. ĐIỀU KHIỂN GIAO DIỆN & MODAL

// ==========================================



/**

 * Mở chế độ Luyện tập theo từng thẻ công thức (Topic 1 - 6)

 * @param {number} topicId 

 */

function openPractice(topicId) {

    isChallenge = false;

    selectedTopics = [topicId];

    score = 0;

    streak = 0;



    const playerEl = document.getElementById('quizPlayer');

    const scoreEl = document.getElementById('scoreDisplay');

    const streakEl = document.getElementById('streakDisplay');



    if (playerEl) playerEl.innerText = `🎯 Luyện tập Thẻ ${topicId}`;

    if (scoreEl) scoreEl.style.display = "none";

    if (streakEl) streakEl.style.display = "none";



    const modal = document.getElementById('quizModal');

    if (modal) modal.classList.add('active');

    

    renderNextQuestion();

}



/**

 * Mở Modal cấu hình Thử thách 20 câu

 */

function openChallengeModal() {

    const modal = document.getElementById('setupModal');

    if (modal) modal.classList.add('active');

}



/**

 * Đóng Modal theo ID

 * @param {string} modalId 

 */

function closeModal(modalId) {

    const modal = document.getElementById(modalId);

    if (modal) modal.classList.remove('active');

}



/**

 * Bắt đầu Thử thách Chuỗi 20 câu

 */

function startChallenge() {

    const nameInput = document.getElementById('studentName');

    const name = (nameInput && nameInput.value.trim()) ? nameInput.value.trim() : "Học sinh";

    

    const checkboxes = document.querySelectorAll('.topic-cb:checked');

    if (checkboxes.length === 0) {

        alert("Vui lòng chọn ít nhất 1 khối công thức để bắt đầu!");

        return;

    }

    

    selectedTopics = Array.from(checkboxes).map(cb => parseInt(cb.value));

    isChallenge = true;

    score = 0;

    streak = 0;



    const playerEl = document.getElementById('quizPlayer');

    const scoreEl = document.getElementById('scoreDisplay');

    const streakEl = document.getElementById('streakDisplay');



    if (playerEl) playerEl.innerText = `👤 ${name}`;

    if (scoreEl) scoreEl.style.display = "inline";

    if (streakEl) streakEl.style.display = "inline";

    

    updateStats();

    closeModal('setupModal');

    

    const quizModal = document.getElementById('quizModal');

    if (quizModal) quizModal.classList.add('active');

    

    renderNextQuestion();

}



/**

 * Cập nhật bảng điểm và chuỗi đúng

 */

function updateStats() {

    const scoreEl = document.getElementById('scoreDisplay');

    const streakEl = document.getElementById('streakDisplay');

    

    if (scoreEl) scoreEl.innerText = `Điểm: ${score}`;

    if (streakEl) streakEl.innerText = `Chuỗi: ${streak}/20 🔥`;



    // Cập nhật kỷ lục cá nhân

    if (score > highScore) {

        highScore = score;

        localStorage.setItem('vector_app_highscore', highScore);

        updateHighScoreDisplay();

    }

}



// ==========================================

// 3. THUẬT TOÁN SINH CÂU HỎI & GÂY NHIỄU SÂU

// ==========================================



const rInt = (min, max) => Math.floor(Math.random() * (max - min + 1)) + min;

const rNonZero = (min, max) => {

    let val = 0;

    while (val === 0) val = rInt(min, max);

    return val;

};



/**

 * Engine sinh câu hỏi trắc nghiệm ngẫu nhiên theo chủ đề được chọn

 * Các phương án nhiễu được tạo từ các bước giải sai phổ biến của học sinh

 */

function generateQuestion() {

    const topic = selectedTopics[Math.floor(Math.random() * selectedTopics.length)];

    let qText = "";

    let correctVal = "";

    let distractors = [];

    let explanation = "";

	let subType = "";


    // ------------------------------------------

    // DẠNG 1: TỌA ĐỘ VECTƠ & ĐIỂM

    // ------------------------------------------

    if (topic === 1) {

        const subType = Math.random() < 0.5 ? 'vector_ab' : 'unit_vector';

        

        if (subType === 'vector_ab') {

            const xA = rInt(-5, 5), yA = rInt(-5, 5);

            const xB = rInt(-5, 5), yB = rInt(-5, 5);

            qText = `Cho hai điểm $A(${xA}; ${yA})$ và $B(${xB}; ${yB})$. Tọa độ của vectơ $\\vec{AB}$ là:`;

            

            const xAns = xB - xA;

            const yAns = yB - yA;

            correctVal = `(${xAns}; ${yAns})`;



            // Gây nhiễu từ sai lầm phổ biến:

            distractors.push(`(${xA - xB}; ${yA - yB})`); // Lỗi 1: Ngược điểm (Lấy A - B)

            distractors.push(`(${xA + xB}; ${yA + yB})`); // Lỗi 2: Nhầm công thức (Cộng tọa độ)

            distractors.push(`(${yAns}; ${xAns})`);       // Lỗi 3: Đáo vị trí hoành - tung



            explanation = `Công thức: $\\vec{AB} = (x_B - x_A; y_B - y_A) = (${xB} - (${xA}); ${yB} - (${yA})) = (${xAns}; ${yAns})$.`;

        } else {

            const x = rNonZero(-6, 6);

            const y = rNonZero(-6, 6);

            qText = `Trong mặt phẳng $Oxy$, cho vectơ $\\vec{u} = ${x}\\vec{i} ${y > 0 ? '+ ' + y : '- ' + Math.abs(y)}\\vec{j}$. Tọa độ của $\\vec{u}$ là:`;

            

            correctVal = `(${x}; ${y})`;

            distractors.push(`(${y}; ${x})`);            // Lỗi 1: Nhầm thứ tự i, j

            distractors.push(`(${-x}; ${-y})`);          // Lỗi 2: Đổi dấu cả 2 tọa độ

            distractors.push(`(${x + y}; ${x - y})`);     // Lỗi 3: Thực hiện phép cộng trừ hệ số



            explanation = `Biểu diễn theo vectơ đơn vị $\\vec{u} = x\\vec{i} + y\\vec{j} \\Rightarrow \\vec{u} = (x; y) = (${x}; ${y})$.`;

        }

    }



    // ------------------------------------------

    // DẠNG 2: ĐỘ DÀI VECTƠ & KHOẢNG CÁCH

    // ------------------------------------------

    else if (topic === 2) {

        const subType = Math.random() < 0.5 ? 'length_u' : 'dist_ab';

        

        if (subType === 'length_u') {

            // Chọn bộ số đẹp Pythagore hoặc số nguyên

            const pyth = [[3, 4], [6, 8], [5, 12]];

            const pair = pyth[Math.floor(Math.random() * pyth.length)];

            const x = pair[0] * (Math.random() < 0.5 ? 1 : -1);

            const y = pair[1] * (Math.random() < 0.5 ? 1 : -1);

            

            qText = `Độ dài của vectơ $\\vec{u} = (${x}; ${y})$ là:`;

            const len = Math.sqrt(x * x + y * y);

            correctVal = `${len}`;



            // Gây nhiễu:

            distractors.push(`${x * x + y * y}`);                   // Lỗi 1: Quên lấy căn bậc hai (x^2 + y^2)

            distractors.push(`${Math.abs(x) + Math.abs(y)}`);       // Lỗi 2: Cộng trực tiếp độ dài hoành + tung

            distractors.push(`${Math.abs(x * x - y * y)}`);         // Lỗi 3: Nhầm dấu trừ trong căn



            explanation = `Công thức độ dài: $|\\vec{u}| = \\sqrt{x^2 + y^2} = \\sqrt{(${x})^2 + (${y})^2} = \\sqrt{${x*x + y*y}} = ${len}$.`;

        } else {

            const xA = rInt(-3, 3), yA = rInt(-3, 3);

            const dx = 3 * (Math.random() < 0.5 ? 1 : -1);

            const dy = 4 * (Math.random() < 0.5 ? 1 : -1);

            const xB = xA + dx, yB = yA + dy;



            qText = `Khoảng cách giữa hai điểm $A(${xA}; ${yA})$ và $B(${xB}; ${yB})$ bằng:`;

            correctVal = `5`;



            distractors.push(`25`);                                 // Lỗi 1: Bỏ quên căn bậc 2

            distractors.push(`${Math.abs(dx) + Math.abs(dy)}`);     // Lỗi 2: Lấy tổng trị tuyệt đối hiệu

            distractors.push(`7`);                                  // Lỗi 3: Tính nhầm căn (3^2+4^2 nhầm sang 3+4)



            explanation = `Khoảng cách $AB = \\sqrt{(x_B-x_A)^2 + (y_B-y_A)^2} = \\sqrt{(${dx})^2 + (${dy})^2} = \\sqrt{25} = 5$.`;

        }

    }



    // ------------------------------------------

    // DẠNG 3: QUAN HỆ HAI VECTƠ (CÙNG PHƯƠNG, BẰNG NHAU)

    // ------------------------------------------

    else if (topic === 3) {

        const subType = Math.random() < 0.5 ? 'collinear_m' : 'equal_vector';

        

        if (subType === 'collinear_m') {

            const k = rNonZero(-3, 4);

            const x1 = rNonZero(-4, 4);

            const y1 = rNonZero(-4, 4);

            const x2 = k * x1;

            

            qText = `Cho $\\vec{a} = (${x1}; ${y1})$ và $\\vec{b} = (${x2}; m)$. Tìm $m$ để hai vectơ $\\vec{a}$ và $\\vec{b}$ cùng phương:`;

            const mCorrect = k * y1;

            correctVal = `m = ${mCorrect}`;



            // Gây nhiễu từ lỗi giải tỉ số x1/x2 = y1/y2:

            distractors.push(`m = ${-mCorrect}`);                   // Lỗi 1: Nhầm dấu hệ số k

            distractors.push(`m = ${y1}`);                          // Lỗi 2: Cho m bằng tung độ ban đầu y1

            distractors.push(`m = ${mCorrect + 2}`);                // Lỗi 3: Lỗi cộng nhầm hệ số khi chuyển vế



            explanation = `Hai vectơ cùng phương khi $\\vec{b} = k\\vec{a} \\Rightarrow \\frac{${x2}}{${x1}} = \\frac{m}{${y1}} = ${k} \\Rightarrow m = ${k} \\cdot (${y1}) = ${mCorrect}$.`;

        } else {

            const xA = rInt(-4, 4), yA = rInt(-4, 4);

            const xB = rInt(-4, 4), yB = rInt(-4, 4);

            const xC = rInt(-4, 4), yC = rInt(-4, 4);

            // Tìm D sao cho ABCD là hình bình hành => AB = DC => xD - xC = xB - xA => xD = xC + xB - xA

            const xD = xC + (xB - xA);

            const yD = yC + (yB - yA);



            qText = `Cho 3 điểm $A(${xA}; ${yA})$, $B(${xB}; ${yB})$, $C(${xC}; ${yC})$. Tọa độ điểm $D$ để $\\vec{AB} = \\vec{DC}$ là:`;

            // AB = DC => xC - xD = xB - xA => xD = xC - (xB - xA)

            const xD_correct = xC - (xB - xA);

            const yD_correct = yC - (yB - yA);

            correctVal = `(${xD_correct}; ${yD_correct})`;



            distractors.push(`(${xD}; ${yD})`);                      // Lỗi 1: Nhầm sang đẳng thức AB = CD (Hình bình hành ABDC)

            distractors.push(`(${xB - xA}; ${yB - yA})`);            // Lỗi 2: Nhầm điểm D chính là vectơ AB

            distractors.push(`(${xD_correct + 1}; ${yD_correct - 1})`);// Lỗi 3: Sai dấu khi chuyển vế



            explanation = `Đẳng thức $\\vec{AB} = \\vec{DC} \\Leftrightarrow (x_B-x_A; y_B-y_A) = (x_C-x_D; y_C-y_D) \\Rightarrow D(${xD_correct}; ${yD_correct})$.`;

        }

    }



    // ------------------------------------------

    // DẠNG 4: PHÉP TOÁN VECTƠ (CỘNG, TRỪ, NHÂN HỆ SỐ)

    // ------------------------------------------

    else if (topic === 4) {

        const x1 = rInt(-4, 4), y1 = rInt(-4, 4);

        const x2 = rInt(-4, 4), y2 = rInt(-4, 4);

        const k = rNonZero(-2, 3);



        if (k === 1) {

            qText = `Cho $\\vec{a}=(${x1}; ${y1})$ và $\\vec{b}=(${x2}; ${y2})$. Tọa độ của $\\vec{u} = \\vec{a} - \\vec{b}$ là:`;

            correctVal = `(${x1 - x2}; ${y1 - y2})`;



            distractors.push(`(${x1 + x2}; ${y1 + y2})`);            // Lỗi 1: Nhầm phép trừ thành phép cộng

            distractors.push(`(${x2 - x1}; ${y2 - y1})`);            // Lỗi 2: Trừ ngược thứ tự b - a

            distractors.push(`(${x1 * x2}; ${y1 * y2})`);            // Lỗi 3: Nhân hai tọa độ với nhau



            explanation = `Phép trừ vectơ: $\\vec{a} - \\vec{b} = (x_1 - x_2; y_1 - y_2) = (${x1}-(${x2}); ${y1}-(${y2})) = (${x1-x2}; ${y1-y2})$.`;

        } else {

            qText = `Cho $\\vec{a}=(${x1}; ${y1})$ và $\\vec{b}=(${x2}; ${y2})$. Tọa độ của $\\vec{v} = ${k}\\vec{a} + \\vec{b}$ là:`;

            const xAns = k * x1 + x2;

            const yAns = k * y1 + y2;

            correctVal = `(${xAns}; ${yAns})`;



            distractors.push(`(${k * x1 + x2}; ${y1 + y2})`);         // Lỗi 1: Chỉ nhân k vào hoành độ, quên nhân tung độ

            distractors.push(`(${k * x1 - x2}; ${k * y1 - y2})`);     // Lỗi 2: Nhầm dấu cộng thành dấu trừ

            distractors.push(`(${x1 + x2}; ${y1 + y2})`);            // Lỗi 3: Quên nhân hệ số k



            explanation = `Tính $k\\vec{a} = (${k*x1}; ${k*y1})$, sau đó cộng với $\\vec{b} \\Rightarrow \\vec{v} = (${xAns}; ${yAns})$.`;

        }

    }



    // ------------------------------------------

    // DẠNG 5: TRUNG ĐIỂM & TRỌNG TÂM

    // ------------------------------------------

    else if (topic === 5) {

        const subType = Math.random() < 0.5 ? 'midpoint' : 'centroid';



        if (subType === 'midpoint') {

            const xA = rInt(-5, 5), yA = rInt(-5, 5);

            const xB = rInt(-5, 5), yB = rInt(-5, 5);

            // Chọn số để tổng chẵn cho tọa độ đẹp

            const xA_c = (xA + xB) % 2 !== 0 ? xA + 1 : xA;

            const yA_c = (yA + yB) % 2 !== 0 ? yA + 1 : yA;



            qText = `Tọa độ trung điểm $M$ của đoạn thẳng $AB$ với $A(${xA_c}; ${yA_c})$ và $B(${xB}; ${yB})$ là:`;

            const xM = (xA_c + xB) / 2;

            const yM = (yA_c + yB) / 2;

            correctVal = `(${xM}; ${yM})`;



            distractors.push(`(${xA_c + xB}; ${yA_c + yB})`);        // Lỗi 1: Quên chia 2 (Tổng đơn thuần)

            distractors.push(`(${xM > 0 ? xM - 1 : xM + 1}; ${yM})`);// Lỗi 2: Lấy hiệu chia 2

            distractors.push(`(${((xA_c + xB)/3).toFixed(1)}; ${((yA_c + yB)/3).toFixed(1)})`); // Lỗi 3: Nhầm chia 3 của trọng tâm



            explanation = `Công thức trung điểm: $M\\left(\\frac{x_A+x_B}{2}; \\frac{y_A+y_B}{2}\\right) = \\left(\\frac{${xA_c}+(${xB})}{2}; \\frac{${yA_c}+(${yB})}{2}\\right) = (${xM}; ${yM})$.`;

        } else {

            const xA = rInt(-4, 4), yA = rInt(-4, 4);

            const xB = rInt(-4, 4), yB = rInt(-4, 4);

            const xC = rInt(-4, 4), yC = rInt(-4, 4);

            

            // Điều chỉnh tổng chia hết cho 3

            const modX = (xA + xB + xC) % 3;

            const modY = (yA + yB + yC) % 3;

            const xA_c = xA - modX;

            const yA_c = yA - modY;



            qText = `Tọa độ trọng tâm $G$ của tam giác $ABC$ với $A(${xA_c}; ${yA_c})$, $B(${xB}; ${yB})$, $C(${xC}; ${yC})$ là:`;

            const xG = (xA_c + xB + xC) / 3;

            const yG = (yA_c + yB + yC) / 3;

            correctVal = `(${xG}; ${yG})`;



            distractors.push(`(${xA_c + xB + xC}; ${yA_c + yB + yC})`); // Lỗi 1: Quên chia 3

            distractors.push(`(${((xA_c + xB + xC)/2).toFixed(1)}; ${((yA_c + yB + yC)/2).toFixed(1)})`); // Lỗi 2: Chia 2 (nhầm trung điểm)

            distractors.push(`(${-xG}; ${-yG})`);                    // Lỗi 3: Nhầm dấu tọa độ trọng tâm



            explanation = `Công thức trọng tâm tam giác: $G\\left(\\frac{x_A+x_B+x_C}{3}; \\frac{y_A+y_B+y_C}{3}\\right) = (${xG}; ${yG})$.`;

        }

    }



    // ------------------------------------------

    // DẠNG 6: TÍCH VÔ HƯỚNG & GÓC

    // ------------------------------------------

    else {

        const subType = Math.random() < 0.5 ? 'dot_product' : 'perpendicular';



        if (subType === 'dot_product') {

            const x1 = rInt(-4, 4), y1 = rInt(-4, 4);

            const x2 = rInt(-4, 4), y2 = rInt(-4, 4);

            qText = `Tích vô hướng $\\vec{a}\\cdot\\vec{b}$ của $\\vec{a}=(${x1}; ${y1})$ và $\\vec{b}=(${x2}; ${y2})$ bằng:`;

            

            const dot = x1 * x2 + y1 * y2;

            correctVal = `${dot}`;



            distractors.push(`${x1 * y2 + x2 * y1}`);               // Lỗi 1: Nhân chéo hoành x tung (x1*y2 + x2*y1)

            distractors.push(`(${x1 * x2}; ${y1 * y2})`);           // Lỗi 2: Nhầm kết quả vô hướng thành cặp tọa độ

            distractors.push(`${x1 * x2 - y1 * y2}`);               // Lỗi 3: Nhầm dấu cộng thành trừ (x1*x2 - y1*y2)



            explanation = `Công thức tích vô hướng: $\\vec{a}\\cdot\\vec{b} = x_1x_2 + y_1y_2 = (${x1})(${x2}) + (${y1})(${y2}) = ${dot}$.`;

        } else {

            const x1 = rNonZero(-3, 4);

            const y1 = rNonZero(-3, 4);

            // Chọn b sao cho a.b = 0 => x1*x2 + y1*y2 = 0 => x2 = y1, y2 = -x1

            const x2 = y1;

            const y2 = -x1;



            qText = `Cho $\\vec{a}=(${x1}; ${y1})$ và $\\vec{b}=(${x2}; m)$. Tìm $m$ để $\\vec{a} \\perp \\vec{b}$:`;

            correctVal = `m = ${y2}`;



            distractors.push(`m = ${-y2}`);                         // Lỗi 1: Nhầm dấu phương trình tích vô hướng

            distractors.push(`m = ${x2}`);                          // Lỗi 2: Cho m bằng hoành độ

            distractors.push(`m = 0`);                              // Lỗi 3: Đoán mò m = 0



            explanation = `Hai vectơ vuông góc $\\vec{a} \\perp \\vec{b} \\Leftrightarrow \\vec{a}\\cdot\\vec{b} = 0 \\Leftrightarrow (${x1})(${x2}) + (${y1})(m) = 0 \\Rightarrow m = ${y2}$.`;

        }

    }



    // Lọc loại bỏ trùng lặp giữa các phương án nhiễu & đáp án đúng

    let uniqueOptions = Array.from(new Set([correctVal, ...distractors]));

    

    // Nếu thiếu lựa chọn, tự động sinh phương án bổ sung hợp lý

    while (uniqueOptions.length < 4) {

        const dummyX = rInt(-8, 8);

        const dummyY = rInt(-8, 8);

        const dummy = correctVal.includes('(') ? `(${dummyX}; ${dummyY})` : `${dummyX}`;

        if (!uniqueOptions.includes(dummy)) {

            uniqueOptions.push(dummy);

        }

    }



    // Trộn ngẫu nhiên 4 phương án

    uniqueOptions = uniqueOptions.sort(() => Math.random() - 0.5);



    return {

        topic: topic,       // Thêm dòng này
		subType: subType,   // Thêm dòng này (ví dụ: 'vector_ab', 'vector_sub'...)
		text: qText,
		correct: correctVal,
		options: uniqueOptions,
		explanation: explanation

    };

}



// ==========================================

// 4. KHỞI TẠO & XỬ LÝ TRẢ LỜI CÂU HỎI

// ==========================================



function renderMathInApp(element) {

    if (window.renderMathInElement && element) {

        renderMathInElement(element, {

            delimiters: [

                {left: '$$', right: '$$', display: true},

                {left: '$', right: '$', display: false},

                {left: '\\(', right: '\\)', display: false},

                {left: '\\[', right: '\\]', display: true}

            ],

            throwOnError: false

        });

    }

}



/**

 * Hiển thị câu hỏi tiếp theo lên màn hình

 */

function renderNextQuestion() {

    const feedback = document.getElementById('feedbackBox');

    const nextBtn = document.getElementById('nextBtn');

    

    if (feedback) {

        feedback.className = "feedback-box";

        feedback.style.display = "none";

    }

    if (nextBtn) nextBtn.style.display = "none";



    currentQuestion = generateQuestion();

    

    const qBox = document.getElementById('questionBox');

    if (qBox) qBox.innerHTML = currentQuestion.text;



    const optionsBox = document.getElementById('optionsBox');

    if (optionsBox) {

        optionsBox.innerHTML = "";

        currentQuestion.options.forEach(opt => {

            const btn = document.createElement('button');

            btn.className = "option-btn";

            btn.innerHTML = `<span>${opt}</span>`;

            btn.onclick = () => checkAnswer(btn, opt);

            optionsBox.appendChild(btn);

        });

    }



    // Render công thức toán KaTeX

    if (window.renderMathInElement) {

        const quizModal = document.getElementById('quizModal');

        if (quizModal) renderMathInApp(quizModal);

    }

}



/**

 * Kiểm tra đáp án khi người dùng bấm chọn

 * @param {HTMLElement} selectedBtn 

 * @param {string} chosenOpt 

 */

function checkAnswer(selectedBtn, chosenOpt) {
    const allBtns = document.querySelectorAll('.option-btn');
    const feedback = document.getElementById('feedbackBox');
    const nextBtn = document.getElementById('nextBtn');

    if (chosenOpt === currentQuestion.correct) {
        // --- ĐÚNG ---
        allBtns.forEach(btn => btn.disabled = true);
        selectedBtn.classList.add('correct');
        triggerHaptic('correct');

        if (feedback) {
            feedback.className = "feedback-box correct";
            feedback.innerHTML = "✅ <b>Chính xác!</b><br>" + currentQuestion.explanation;
        }

        if (isChallenge) {
            score += 1;
            streak += 1;
            updateStats();
            if (streak >= 20) {
                triggerHaptic('win');
                alert(`🎉 CHÚC MỪNG! Bạn đã xuất sắc vượt qua Thử Thách 20 câu đúng liên tiếp!\n\nTổng điểm đạt được: ${score}`);
                closeModal('quizModal');
                return;
            }
        }
        if (nextBtn) nextBtn.style.display = "block";

    } else {
        // --- SAI ---
        triggerHaptic('wrong');

        if (!isChallenge) {
            // 📌 CHẾ ĐỘ LUYỆN TẬP: Chỉ khóa nút vừa chọn sai, hiện bảng nhắc công thức từ Excel cho làm lại
            console.log("chọn sai");
			selectedBtn.disabled = true; 
            selectedBtn.classList.add('wrong');

            // Tạo khóa tra cứu từ topic và subType của câu hỏi hiện tại
            const key = `${currentQuestion.topic}_${currentQuestion.subType}`;
            const theory = excelTheoryData[key];
            
            let hintHtml = "❌ <b>Chưa chính xác, hãy xem lại công thức và thử lại nhé!</b><br>";
            if (theory) {
                hintHtml += `<div style="margin-top: 8px; padding-top: 8px; border-top: 1px dashed #fca5a5; font-size: 0.9rem;">
                    📌 <b>Công thức chuẩn:</b> ${theory.formula}<br>
                    💡 <b>Hướng dẫn giải:</b> ${theory.guide}<br>
                    ⚠️ <b>Lưu ý lỗi sai:</b> ${theory.mistake}
                </div>`;
            } else {
                hintHtml += `<br>${currentQuestion.explanation}`;
            }

            if (feedback) {
                feedback.className = "feedback-box wrong";
				feedback.style.display = "block";
                feedback.innerHTML = hintHtml;
            }
            // Không hiện nút "Câu tiếp theo" để học sinh tiếp tục chọn các đáp án khác cho đến khi đúng

        } else {
            // 📌 CHẾ ĐỘ THỬ THÁCH: Khóa tất cả, trừ điểm, reset chuỗi
            allBtns.forEach(btn => {
                btn.disabled = true;
                if (btn.innerText.trim() === currentQuestion.correct.trim()) {
                    btn.classList.add('correct');
                }
            });
            selectedBtn.classList.add('wrong');

            if (feedback) {
                feedback.className = "feedback-box wrong";
                feedback.innerHTML = "❌ <b>Rất tiếc, bạn đã chọn sai!</b> -2 điểm (Chuỗi reset về 0).<br>" + currentQuestion.explanation;
            }

            score -= 2;
            streak = 0;
            updateStats();
            if (nextBtn) nextBtn.style.display = "block";
        }
    }

    // Render lại công thức toán KaTeX cho đẹp
    if (window.renderMathInElement && feedback) {
        renderMathInElement(feedback, {
            delimiters: [
                {left: '$$', right: '$$', display: true},
                {left: '$', right: '$', display: false},
                {left: '\\(', right: '\\)', display: false},
                {left: '\\[', right: '\\]', display: true}
            ],
            throwOnError: false
        });
    }
}



// ==========================================

// 5. CHẾ ĐỘ TOÀN MÀN HÌNH (FULLSCREEN API)

// ==========================================



function toggleFullScreen() {

    const docEl = document.documentElement;

    if (!document.fullscreenElement && !document.webkitFullscreenElement) {

        if (docEl.requestFullscreen) {

            docEl.requestFullscreen();

        } else if (docEl.webkitRequestFullscreen) { // Hỗ trợ Safari/iOS

            docEl.webkitRequestFullscreen();

        }

    } else {

        if (document.exitFullscreen) {

            document.exitFullscreen();

        } else if (document.webkitExitFullscreen) {

            document.webkitExitFullscreen();

        }

    }

}



// Cập nhật biểu tượng nút khi thay đổi trạng thái màn hình

['fullscreenchange', 'webkitfullscreenchange'].forEach(evt => {

    document.addEventListener(evt, () => {

        const icon = document.getElementById('fsIcon');

        if (icon) {

            const isFS = document.fullscreenElement || document.webkitFullscreenElement;

            icon.innerText = isFS ? '✖' : '⛶';

        }

    });

});

// --- GÁN CÁC HÀM VÀO WINDOW ĐỂ HTML GỌI ĐƯỢC KHI DÙNG TYPE="MODULE" ---
window.openPractice = openPractice;
window.openChallengeModal = openChallengeModal;
window.closeModal = closeModal;
window.startChallenge = startChallenge;
window.renderNextQuestion = renderNextQuestion;
window.toggleFullScreen = toggleFullScreen;