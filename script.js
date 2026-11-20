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

// 1. Hàm hỗ trợ rút gọn căn bậc hai (VD: sqrt(12) -> 2\sqrt{3}, sqrt(25) -> 5, sqrt(13) -> \sqrt{13})
function simplifyRadical(n) {
    if (n <= 0) return "$0$";
    let a = 1;
    for (let i = Math.floor(Math.sqrt(n)); i >= 2; i--) {
        if (n % (i * i) === 0) {
            a = i;
            break;
        }
    }
    let b = n / (a * a);
    if (b === 1) return `$${a}$`;
    if (a === 1) return `$\\sqrt{${b}}$`;
    return `$${a}\\sqrt{${b}}$`;
}

/**
 * 1. HÀM TÌM ƯỚC CHUNG LỚN NHẤT (GCD) ĐỂ RÚT GỌN PHÂN SỐ
 */
function gcd(a, b) {
    a = Math.abs(a);
    b = Math.abs(b);
    while (b) {
        let t = b;
        b = a % b;
        a = t;
    }
    return a;
}

/**
 * 2. HÀM ĐỊNH DẠNG PHÂN SỐ CHUẨN LATEX KATEX
 * Tự động chuyển đổi cặp (tử số, mẫu số) thành chuỗi phân số LaTeX:
 * VD: (6, 2) -> "3", (-7, 2) -> "-\\frac{7}{2}", (0, 3) -> "0"
 */
function formatFrac(num, den) {
    if (num === 0) return "0";
    if (den < 0) {
        num = -num;
        den = -den;
    }
    const g = gcd(num, den);
    num = num / g;
    den = den / g;
    if (den === 1) return `${num}`;
    if (num < 0) return `-\\frac{${Math.abs(num)}}{${den}}`;
    return `\\frac{${num}}{${den}}`;
}

function simplifyFraction(num, den) {
    if (den < 0) {
        num = -num;
        den = -den;
    }
    let g = gcd(num, den);
    num = num / g;
    den = den / g;
    if (den === 1) return `$${num}$`;
    return `$\\frac{${num}}{${den}}$`;
}

function simplifyRadicalFraction(dot, sum1, sum2) {
    if (dot === 0) return "$0$";
    let prod = sum1 * sum2;
    let a = 1;
    for (let i = Math.floor(Math.sqrt(prod)); i >= 2; i--) {
        if (prod % (i * i) === 0) {
            a = i;
            break;
        }
    }
    let b = prod / (a * a);
    let num = dot;
    let den = a * b;
    if (den < 0) {
        num = -num;
        den = -den;
    }
    let g = gcd(num, den);
    num = num / g;
    den = den / g;

    let sign = num < 0 ? "-" : "";
    let absNum = Math.abs(num);

    let res = "";
    if (b === 1) {
        if (den === 1) res = `${num}`;
        else res = `\\frac{${num}}{${den}}`;
    } else {
        let sqrtPart = (absNum === 1) ? `\\sqrt{${b}}` : `${absNum}\\sqrt{${b}}`;
        if (den === 1) res = `${sign}${sqrtPart}`;
        else res = `${sign}\\frac{${sqrtPart}}{${den}}`;
    }

    // Bọc trong cặp dấu $...$ để KaTeX tự động render trên giao diện Mobile
    return `$${res}$`;
}


/**
 * 3. HÀM TẠO CHUỖI TỌA ĐỘ BỌC KATEX CHUẨN (TỰ ĐỘNG NỞ NGOẶC NẾU CÓ PHÂN SỐ)
 * VD: formatPoint("3", "\\frac{5}{2}") -> "$\\left(3; \\frac{5}{2}\\right)$"
 */
function formatPoint(xStr, yStr) {
    return `$\\left(${xStr}; ${yStr}\\right)$`;
}


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
			// Sinh ngẫu nhiên hoành độ và tung độ (từ -6 đến 6, khác 0)
			let x = rInt(-6, 6);
			let y = rInt(-6, 6);
			if (x === 0 && y === 0) x = 3; // Đảm bảo không trùng gốc tọa độ

			const sumSq = x * x + y * y;
			correctVal = simplifyRadical(sumSq);

			qText = `Độ dài của vectơ $\\vec{u} = (${x}; ${y})$ là:`;

			// GÂY NHIỄU THEO CÁC LỖI SAI PHỔ BIẾN CỦA HỌC SINH:
			// Lỗi 1: Quên lấy căn bậc hai (chỉ tính x^2 + y^2)
			const d1 = `${sumSq}`;
			
			// Lỗi 2: Cộng trực tiếp độ dài tuyệt đối |x| + |y| (bỏ qua bình phương và căn)
			const d2 = `${Math.abs(x) + Math.abs(y)}`;
			
			// Lỗi 3: Nhầm dấu trừ trong căn \sqrt{|x^2 - y^2|}
			const diffSq = Math.abs(x * x - y * y);
			const d3 = simplifyRadical(diffSq);

			distractors = [d1, d2, d3];

			// Lời giải chi tiết hỗ trợ tự học
			const radStr = (Math.sqrt(sumSq) % 1 === 0) ? `${Math.sqrt(sumSq)}` : `\\sqrt{${sumSq}}`;
			const finalStr = (correctVal !== radStr) ? `= ${correctVal}` : '';
			explanation = `Công thức độ dài: $|\\vec{u}| = \\sqrt{x^2 + y^2} = \\sqrt{(${x})^2 + (${y})^2} = \\sqrt{${sumSq}} $${finalStr}.`;

		} else {
			// Sinh ngẫu nhiên tọa độ 2 điểm A và B
			const xA = rInt(-5, 5), yA = rInt(-5, 5);
			let dx = rInt(-5, 5), dy = rInt(-5, 5);
			if (dx === 0 && dy === 0) dx = 4; // Đảm bảo A và B không trùng nhau

			const xB = xA + dx;
			const yB = yA + dy;

			const sumSq = dx * dx + dy * dy;
			correctVal = simplifyRadical(sumSq);

			qText = `Khoảng cách giữa hai điểm $A(${xA}; ${yA})$ và $B(${xB}; ${yB})$ bằng:`;

			// GÂY NHIỄU THEO CÁC LỖI SAI PHỔ BIẾN CỦA HỌC SINH:
			// Lỗi 1: Quên khai căn (tính (xB-xA)^2 + (yB-yA)^2)
			const d1 = `${sumSq}`;
			
			// Lỗi 2: Tổng trị tuyệt đối hiệu tọa độ |xB-xA| + |yB-yA|
			const d2 = `${Math.abs(dx) + Math.abs(dy)}`;
			
			// Lỗi 3: Nhầm dấu trừ trong công thức căn \sqrt{|(xB-xA)^2 - (yB-yA)^2|}
			const diffSq = Math.abs(dx * dx - dy * dy);
			const d3 = simplifyRadical(diffSq);

			distractors = [d1, d2, d3];

			// Lời giải chi tiết
			const radStr = (Math.sqrt(sumSq) % 1 === 0) ? `${Math.sqrt(sumSq)}` : `\\sqrt{${sumSq}}`;
			const finalStr = (correctVal !== radStr) ? `= ${correctVal}` : '';
			explanation = `Khoảng cách $AB = \\sqrt{(x_B-x_A)^2 + (y_B-y_A)^2} = \\sqrt{(${dx})^2 + (${dy})^2} = \\sqrt{${sumSq}} $${finalStr}.`;
		}

		// Đảm bảo các phương án nhiễu luôn duy nhất và không trùng đáp án đúng
		let uniqueDistractors = [];
		for (let d of distractors) {
			if (d !== correctVal && !uniqueDistractors.includes(d)) {
				uniqueDistractors.push(d);
			}
		}
		// Nếu bị trùng làm thiếu đáp án, bổ sung các phương án hợp lý
		let backupVal = 1;
		while (uniqueDistractors.length < 3) {
			let candidate = `${backupVal}`;
			if (candidate !== correctVal && !uniqueDistractors.includes(candidate)) {
				uniqueDistractors.push(candidate);
			}
			backupVal++;
		}
		distractors = uniqueDistractors;
	}



    // ------------------------------------------

    // DẠNG 3: QUAN HỆ HAI VECTƠ (CÙNG PHƯƠNG, BẰNG NHAU)

    // ------------------------------------------

    else if (topic === 3) {
		distractors = []; // Reset danh sách phương án nhiễu
		const rand = Math.random();

		if (rand < 0.6) {
			// =================================================================
			// LOẠI 1 (60%): NHẬN BIẾT HAI VECTƠ CÙNG PHƯƠNG / CÙNG HƯỚNG / NGƯỢC HƯỚNG
			// Mẫu câu hỏi: "Cặp vectơ nào sau đây cùng phương / cùng hướng / ngược hướng?"
			// =================================================================
			const reqTypes = ['cùng phương', 'cùng hướng', 'ngược hướng'];
			const req = reqTypes[Math.floor(Math.random() * reqTypes.length)];

			qText = `Cặp vectơ nào sau đây ${req}?`;

			// Sinh cặp vectơ đúng tương ứng với yêu cầu
			let u_x = rNonZero(-4, 4);
			let u_y = rNonZero(-4, 4);

			let k_correct;
			if (req === 'cùng hướng') {
				k_correct = rInt(2, 4); // k > 0
			} else if (req === 'ngược hướng') {
				k_correct = rInt(-4, -2); // k < 0
			} else { // cùng phương
				const listK = [-3, -2, 2, 3];
				k_correct = listK[Math.floor(Math.random() * listK.length)];
			}

			let v_x = k_correct * u_x;
			let v_y = k_correct * u_y;

			correctVal = `$\\vec{u} = (${u_x}; ${u_y})$ và $\\vec{v} = (${v_x}; ${v_y})$`;

			// GÂY NHIỄU THEO BẪY TƯ DUY GDPT 2018:
			// Bẫy 1: Cặp vectơ có tính chất hướng ngược lại (Ví dụ: Hỏi cùng hướng thì đưa ra cặp ngược hướng)
			let k_opposite = (req === 'cùng hướng') ? rInt(-3, -2) : (req === 'ngược hướng' ? rInt(2, 3) : 2);
			let d1_ux = rNonZero(-4, 4), d1_uy = rNonZero(-4, 4);
			let d1_vx = k_opposite * d1_ux, d1_vy = k_opposite * d1_uy;
			let d1 = `$\\vec{u} = (${d1_ux}; ${d1_uy})$ và $\\vec{v} = (${d1_vx}; ${d1_vy})$`;

			// Bẫy 2: Cặp vectơ không cùng phương (Tỉ số x1/x2 != y1/y2)
			let d2_ux = rNonZero(-4, 4), d2_uy = rNonZero(-4, 4);
			let d2_vx = d2_ux * 2, d2_vy = d2_uy * 3 + 1; // Sai tỉ lệ hoành/tung
			let d2 = `$\\vec{u} = (${d2_ux}; ${d2_uy})$ và $\\vec{v} = (${d2_vx}; ${d2_vy})$`;

			// Bẫy 3: Cặp vectơ vuông góc (u.v = 0, tích hoành + tích tung = 0)
			let d3_ux = rNonZero(-4, 4), d3_uy = rNonZero(-4, 4);
			let d3_vx = -d3_uy, d3_vy = d3_ux; 
			let d3 = `$\\vec{u} = (${d3_ux}; ${d3_uy})$ và $\\vec{v} = (${d3_vx}; ${d3_vy})$`;

			distractors.push(d1, d2, d3);

			if (req === 'cùng phương') {
				explanation = `Cặp vectơ $\\vec{u} = (${u_x}; ${u_y})$ và $\\vec{v} = (${v_x}; ${v_y})$ cùng phương vì $\\vec{v} = ${k_correct}\\vec{u}$ (tỉ số $\\frac{${v_x}}{${u_x}} = \\frac{${v_y}}{${u_y}} = ${k_correct}$).`;
			} else if (req === 'cùng hướng') {
				explanation = `Cặp vectơ $\\vec{u} = (${u_x}; ${u_y})$ và $\\vec{v} = (${v_x}; ${v_y})$ cùng hướng vì $\\vec{v} = ${k_correct}\\vec{u}$ với hệ số tỷ lệ $k = ${k_correct} > 0$.`;
			} else {
				explanation = `Cặp vectơ $\\vec{u} = (${u_x}; ${u_y})$ và $\\vec{v} = (${v_x}; ${v_y})$ ngược hướng vì $\\vec{v} = ${k_correct}\\vec{u}$ với hệ số tỷ lệ $k = ${k_correct} < 0$.`;
			}

		} else if (rand < 0.8) {
			// =================================================================
			// LOẠI 2 (20%): TÌM GIÁ TRỊ BIẾN ĐỂ HAI VECTƠ CÙNG PHƯƠNG / CÙNG HƯỚNG / NGƯỢC HƯỚNG
			// =================================================================
			const mode = rInt(0, 2); // 0: cùng phương, 1: cùng hướng, 2: ngược hướng
			const x1 = rNonZero(-4, 4);
			const y1 = rNonZero(-4, 4);

			let k = 2;
			if (mode === 1) k = rInt(2, 4);        // Cùng hướng -> k > 0
			else if (mode === 2) k = rInt(-4, -2); // Ngược hướng -> k < 0
			else {
				const listK = [-3, -2, 2, 3];
				k = listK[Math.floor(Math.random() * listK.length)];
			}

			const x2 = k * x1;
			const mCorrect = k * y1;

			let relStr = (mode === 0) ? "cùng phương" : (mode === 1 ? "cùng hướng" : "ngược hướng");

			qText = `Cho hai vectơ $\\vec{a} = (${x1}; ${y1})$ và $\\vec{b} = (${x2}; m)$. Tìm $m$ để $\\vec{a}$ và $\\vec{b}$ ${relStr}:`;
			correctVal = `m = ${mCorrect}`;

			// GÂY NHIỄU TRỰC TIẾP TỪ PHƯƠNG PHÁP GIẢI:
			// Nhiễu 1: Sai dấu tích chéo
			distractors.push(`m = ${-mCorrect}`);
			// Nhiễu 2: Cho m bằng tung độ y1 ban đầu (bỏ qua k)
			distractors.push(`m = ${y1}`);
			// Nhiễu 3: Sai lầm biến đổi cộng trừ hệ số
			let d3_val = (mCorrect > 0) ? mCorrect + 2 : mCorrect - 2;
			if (d3_val === y1 || d3_val === -mCorrect) d3_val += 3;
			distractors.push(`m = ${d3_val}`);

			if (mode === 0) {
				explanation = `Hai vectơ cùng phương khi $\\frac{${x2}}{${x1}} = \\frac{m}{${y1}} = ${k} \\Rightarrow m = ${k} \\cdot (${y1}) = ${mCorrect}$.`;
			} else if (mode === 1) {
				explanation = `Tỉ số hoành độ $\\frac{x_2}{x_1} = \\frac{${x2}}{${x1}} = ${k} > 0$. Để hai vectơ cùng hướng thì $m = ${k} \\cdot (${y1}) = ${mCorrect}$.`;
			} else {
				explanation = `Tỉ số hoành độ $\\frac{x_2}{x_1} = \\frac{${x2}}{${x1}} = ${k} < 0$. Để hai vectơ ngược hướng thì $m = ${k} \\cdot (${y1}) = ${mCorrect}$.`;
			}


		} else {
			// =================================================================
			// LOẠI 3 (20%): TÌM TỌA ĐỘ ĐIỂM D ĐỂ VEC(AB) = VEC(DC)
			// =================================================================
			let xA = rInt(-4, 4), yA = rInt(-4, 4);
			let xB = rInt(-4, 4), yB = rInt(-4, 4);
			let xC = rInt(-4, 4), yC = rInt(-4, 4);

			// Đảm bảo 3 điểm không thẳng hàng
			while ((xB === xA && yB === yA) || (xC === xA && yC === yA) || ((xB - xA) * (yC - yA) === (yB - yA) * (xC - xA))) {
				xB = rInt(-4, 4); yB = rInt(-4, 4);
				xC = rInt(-4, 4); yC = rInt(-4, 4);
			}

			// Tọa độ D thỏa mãn vec(AB) = vec(DC) => xC - xD = xB - xA => xD = xC - (xB - xA)
			const xD_correct = xC - (xB - xA);
			const yD_correct = yC - (yB - yA);

			qText = `Cho ba điểm $A(${xA}; ${yA})$, $B(${xB}; ${yB})$, $C(${xC}; ${yC})$. Tọa độ điểm $D$ để $\\vec{AB} = \\vec{DC}$ (tứ giác $ABCD$ là hình bình hành) là:`;
			correctVal = `D(${xD_correct}; ${yD_correct})`;

			// GÂY NHIỄU BÁM SÁT CÁC LỖI SAI PHỔ BIẾN CỦA HỌC SINH:
			// Lỗi 1: Nhầm sang đẳng thức vec(AB) = vec(CD) (Hình bình hành ABDC) => xD = xC + (xB - xA)
			const xD_err1 = xC + (xB - xA);
			const yD_err1 = yC + (yB - yA);
			distractors.push(`D(${xD_err1}; ${yD_err1})`);

			// Lỗi 2: Nhầm điểm D chính là tọa độ của vectơ vec(AB)
			distractors.push(`D(${xB - xA}; ${yB - yA})`);

			// Lỗi 3: Sai dấu khi chuyển vế (-xD_correct; -yD_correct)
			distractors.push(`D(${-xD_correct}; ${-yD_correct})`);

			explanation = `Đẳng thức $\\vec{AB} = \\vec{DC}$ <br>` + `$\\Leftrightarrow (${xB - xA}; ${yB - yA}) = (${xC} - x_D; ${yC} - y_D)$<br>` +
						  `$\\Rightarrow \\begin{cases} x_D = ${xC} - (${xB - xA}) = ${xD_correct} \\\\ y_D = ${yC} - (${yB - yA}) = ${yD_correct} \\end{cases} \\Rightarrow D(${xD_correct}; ${yD_correct})$.`;
					}

		// ĐẢM BẢO CÁC PHƯƠNG ÁN NHIỄU KHÔNG TRÙNG ĐÁP ÁN ĐÚNG VÀ DUY NHẤT
		let uniqueDistractors = [];
		for (let d of distractors) {
			if (d !== correctVal && !uniqueDistractors.includes(d)) {
				uniqueDistractors.push(d);
			}
		}
		let backupVal = 1;
		while (uniqueDistractors.length < 3) {
				let candidate = formatPoint(`${backupVal}`, `${backupVal + 1}`);
				if (candidate !== correctVal && !uniqueDistractors.includes(candidate)) {
					uniqueDistractors.push(candidate);
				}
				backupVal++;
		}
		distractors = uniqueDistractors;
	}




    // ------------------------------------------

    // DẠNG 4: PHÉP TOÁN VECTƠ (CỘNG, TRỪ, NHÂN HỆ SỐ)

    // ------------------------------------------

    // =================================================================
	// DẠNG 4: CÁC PHÉP TOÁN VECTƠ (GDPT 2018 - TOÁN 10)
	// TỶ LỆ PHÂN BỔ: 2 : 2 : 2 : 4 (20% Tổng | 20% Hiệu | 20% k.a | 40% k.a + m.b)
	// =================================================================

	else if (topic === 4 || topic === 'vector_op') {
		distractors = []; // Reset danh sách phương án nhiễu
		const rand = Math.random();

		if (rand < 0.2) {
			// =================================================================
			// DẠNG 4.1 (20%): VECTƠ TỔNG (u = a + b)
			// =================================================================
			const x1 = rInt(-5, 5), y1 = rInt(-5, 5);
			const x2 = rInt(-5, 5), y2 = rInt(-5, 5);

			qText = `Cho $\\vec{a}=(${x1}; ${y1})$ và $\\vec{b}=(${x2}; ${y2})$. Tọa độ của vectơ tổng $\\vec{u} = \\vec{a} + \\vec{b}$ là:`;
			const xAns = x1 + x2;
			const yAns = y1 + y2;
			correctVal = `(${xAns}; ${yAns})`;

			// GÂY NHIỄU TỪ QUÁ TRÌNH GIẢI TOÁN CỦA HỌC SINH:
			distractors.push(`(${x1 - x2}; ${y1 - y2})`);             // Lỗi 1: Nhầm phép cộng thành phép trừ
			distractors.push(`(${x1 * x2}; ${y1 * y2})`);             // Lỗi 2: Nhân hoành với hoành, tung với tung
			distractors.push(`(${x1 + y2}; ${y1 + x2})`);             // Lỗi 3: Nhầm lẫn giữa hoành độ và tung độ

			explanation = `Phép cộng vectơ: $\\vec{a} + \\vec{b} = (x_1 + x_2; y_1 + y_2) = (${x1} + (${x2}); ${y1} + (${y2})) = (${xAns}; ${yAns})$.`;

		} else if (rand < 0.4) {
			// =================================================================
			// DẠNG 4.2 (20%): VECTƠ HIỆU (u = a - b)
			// =================================================================
			const x1 = rInt(-5, 5), y1 = rInt(-5, 5);
			const x2 = rInt(-5, 5), y2 = rInt(-5, 5);

			qText = `Cho $\\vec{a}=(${x1}; ${y1})$ và $\\vec{b}=(${x2}; ${y2})$. Tọa độ của vectơ hiệu $\\vec{u} = \\vec{a} - \\vec{b}$ là:`;
			const xAns = x1 - x2;
			const yAns = y1 - y2;
			correctVal = `(${xAns}; ${yAns})`;

			// GÂY NHIỄU TỪ QUÁ TRÌNH GIẢI TOÁN CỦA HỌC SINH:
			distractors.push(`(${x1 + x2}; ${y1 + y2})`);             // Lỗi 1: Nhầm phép trừ thành phép cộng
			distractors.push(`(${x2 - x1}; ${y2 - y1})`);             // Lỗi 2: Trừ ngược thứ tự b - a
			distractors.push(`(${x1 * x2}; ${y1 * y2})`);             // Lỗi 3: Nhân hai tọa độ với nhau

			explanation = `Phép trừ vectơ: $\\vec{a} - \\vec{b} = (x_1 - x_2; y_1 - y_2) = (${x1} - (${x2}); ${y1} - (${y2})) = (${xAns}; ${yAns})$.`;

		} else if (rand < 0.6) {
			// =================================================================
			// DẠNG 4.3 (20%): TÍCH VECTƠ VỚI MỘT SỐ (u = k.a)
			// =================================================================
			const x1 = rNonZero(-4, 4), y1 = rNonZero(-4, 4);
			let k = rNonZero(-3, 4);
			if (k === 1) k = -2; // Tránh k = 1 gây trùng đề bài

			qText = `Cho $\\vec{a}=(${x1}; ${y1})$. Tọa độ của vectơ $\\vec{u} = ${k}\\vec{a}$ là:`;
			const xAns = k * x1;
			const yAns = k * y1;
			correctVal = `(${xAns}; ${yAns})`;

			// GÂY NHIỄU TỪ QUÁ TRÌNH GIẢI TOÁN CỦA HỌC SINH:
			distractors.push(`(${k * x1}; ${y1})`);                   // Lỗi 1: Chỉ nhân k vào hoành độ, quên nhân tung độ
			distractors.push(`(${x1}; ${k * y1})`);                   // Lỗi 2: Chỉ nhân k vào tung độ, quên nhân hoành độ
			distractors.push(`(${x1 + k}; ${y1 + k})`);               // Lỗi 3: Cộng k vào tọa độ thay vì nhân

			explanation = `Phép nhân vectơ với một số: $k\\vec{a} = (k \\cdot x_1; k \\cdot y_1) = (${k} \\cdot (${x1}); ${k} \\cdot (${y1})) = (${xAns}; ${yAns})$.`;

		} else {
			// =================================================================
			// DẠNG 4.4 (40%): TỔ HỢP TUYẾN TÍNH VECTƠ (v = k.a + m.b)
			// =================================================================
			const x1 = rInt(-4, 4), y1 = rInt(-4, 4);
			const x2 = rInt(-4, 4), y2 = rInt(-4, 4);
			let k = rNonZero(-3, 3);
			let m = rNonZero(-3, 3);
			if (k === 1 && m === 1) m = -2; // Tránh k=1, m=1 trùng với phép cộng đơn thuần

			// Định dạng biểu thức KaTeX hiển thị đẹp mắt (VD: 2\vec{a} - 3\vec{b})
			let kPart = (k === 1) ? "\\vec{a}" : (k === -1 ? "-\\vec{a}" : `${k}\\vec{a}`);
			let mPart = "";
			if (m === 1) mPart = " + \\vec{b}";
			else if (m === -1) mPart = " - \\vec{b}";
			else if (m > 0) mPart = ` + ${m}\\vec{b}`;
			else mPart = ` - ${Math.abs(m)}\\vec{b}`;
			let exprStr = kPart + mPart;

			qText = `Cho $\\vec{a}=(${x1}; ${y1})$ và $\\vec{b}=(${x2}; ${y2})$. Tọa độ của vectơ $\\vec{v} = ${exprStr}$ là:`;
			const xAns = k * x1 + m * x2;
			const yAns = k * y1 + m * y2;
			correctVal = `(${xAns}; ${yAns})`;

			// GÂY NHIỄU TỪ QUÁ TRÌNH GIẢI TOÁN CỦA HỌC SINH:
			distractors.push(`(${k * x1 + x2}; ${k * y1 + y2})`);     // Lỗi 1: Quên nhân hệ số m vào vectơ b
			distractors.push(`(${k * x1 - m * x2}; ${k * y1 - m * y2})`); // Lỗi 2: Sai dấu phép tính giữa 2 tích
			distractors.push(`(${k * x1 + m * x2}; ${y1 + y2})`);     // Lỗi 3: Quên nhân hệ số vào tung độ

			explanation = `Ta có $k\\vec{a} = (${k * x1}; ${k * y1})$ và $m\\vec{b} = (${m * x2}; ${m * y2})$.<br>` +
						  `$\\Rightarrow \\vec{v} = (${k * x1} + (${m * x2}); ${k * y1} + (${m * y2})) = (${xAns}; ${yAns})$.`;
		}

		// =================================================================
		// ĐẢM BẢO CÁC PHƯƠNG ÁN NHIỄU DUY NHẤT VÀ KHÔNG TRÙNG ĐÁP ÁN ĐÚNG
		// =================================================================
		let uniqueDistractors = [];
		for (let d of distractors) {
			if (d !== correctVal && !uniqueDistractors.includes(d)) {
				uniqueDistractors.push(d);
			}
		}

		let backupVal = 1;
		while (uniqueDistractors.length < 3) {
			let candidate = `(${backupVal}; ${backupVal + 1})`;
			if (candidate !== correctVal && !uniqueDistractors.includes(candidate)) {
				uniqueDistractors.push(candidate);
			}
			backupVal++;
		}

		distractors = uniqueDistractors;
	}



    // ------------------------------------------

    // DẠNG 5: TRUNG ĐIỂM & TRỌNG TÂM

    // ------------------------------------------

    else if (topic === 5) {
		const subType = Math.random() < 0.5 ? 'midpoint' : 'centroid';

		if (subType === 'midpoint') {
			// Sinh ngẫu nhiên tọa độ hai điểm A và B (không ép chẵn/lẻ)
			const xA = rInt(-6, 6), yA = rInt(-6, 6);
			const xB = rInt(-6, 6), yB = rInt(-6, 6);

			qText = `Tọa độ trung điểm $M$ của đoạn thẳng $AB$ với $A(${xA}; ${yA})$ và $B(${xB}; ${yB})$ là:`;

			const sumX = xA + xB;
			const sumY = yA + yB;
			const diffX = xA - xB;
			const diffY = yA - yB;

			const xM = formatFrac(sumX, 2);
			const yM = formatFrac(sumY, 2);
			correctVal = formatPoint(xM, yM);

			// BẪY TRẮC NGHIỆM TỪ NGUYÊN NHÂN LỖI SAI CỦA HỌC SINH:
			const d1 = formatPoint(`${sumX}`, `${sumY}`);                           // Lỗi 1: Quên chia 2 (Tổng đơn thuần)
			const d2 = formatPoint(formatFrac(diffX, 2), formatFrac(diffY, 2));     // Lỗi 2: Nhầm phép trừ (Hiệu chia 2)
			const d3 = formatPoint(formatFrac(sumX, 3), formatFrac(sumY, 3));       // Lỗi 3: Nhầm chia 3 (Trọng tâm)

			distractors = [d1, d2, d3];

			// Lời giải chi tiết
			explanation = `Công thức trung điểm đoạn thẳng $AB$: $M\\left(\\frac{x_A+x_B}{2}; \\frac{y_A+y_B}{2}\\right) \\Rightarrow M\\left(\\frac{${xA}+(${xB})}{2}; \\frac{${yA}+(${yB})}{2}\\right) \\Rightarrow M $${correctVal}.`;

		} else {
			// Sinh ngẫu nhiên tọa độ 3 đỉnh A, B, C của tam giác
			const xA = rInt(-5, 5), yA = rInt(-5, 5);
			const xB = rInt(-5, 5), yB = rInt(-5, 5);
			const xC = rInt(-5, 5), yC = rInt(-5, 5);

			qText = `Tọa độ trọng tâm $G$ của tam giác $ABC$ với $A(${xA}; ${yA})$, $B(${xB}; ${yB})$, $C(${xC}; ${yC})$ là:`;

			const sumX = xA + xB + xC;
			const sumY = yA + yB + yC;

			const xG = formatFrac(sumX, 3);
			const yG = formatFrac(sumY, 3);
			correctVal = formatPoint(xG, yG);

			// BẪY TRẮC NGHIỆM TỪ NGUYÊN NHÂN LỖI SAI CỦA HỌC SINH:
			const d1 = formatPoint(`${sumX}`, `${sumY}`);                           // Lỗi 1: Quên chia 3 (Tổng hoành/tung độ)
			const d2 = formatPoint(formatFrac(sumX, 2), formatFrac(sumY, 2));       // Lỗi 2: Chia 2 (Nhầm công thức trung điểm)
			const d3 = formatPoint(formatFrac(-sumX, 3), formatFrac(-sumY, 3));     // Lỗi 3: Nhầm dấu tọa độ (-x_G; -y_G)

			distractors = [d1, d2, d3];

			// Lời giải chi tiết
			explanation = `Công thức trọng tâm tam giác $ABC$: $G\\left(\\frac{x_A+x_B+x_C}{3}; \\frac{y_A+y_B+y_C}{3}\\right) \\Rightarrow G \\left(\\frac{${xA}+(${xB})+(${xC})}{3}; \\frac{${yA}+(${yB})+(${yC})}{3}\\right) \\Rightarrow G$${correctVal}.`;
		}

		// LỌC ĐẢM BẢO CÁC PHƯƠNG ÁN NHIỄU DUY NHẤT VÀ KHÔNG TRÙNG ĐÁP ÁN ĐÚNG
		let uniqueDistractors = [];
		for (let d of distractors) {
			if (d !== correctVal && !uniqueDistractors.includes(d)) {
				uniqueDistractors.push(d);
			}
		}

		// Bổ sung phương án dự phòng nếu trùng
		let backupVal = 1;
		while (uniqueDistractors.length < 3) {
			let candidate = formatPoint(`${backupVal}`, `${backupVal + 1}`);
			if (candidate !== correctVal && !uniqueDistractors.includes(candidate)) {
				uniqueDistractors.push(candidate);
			}
			backupVal++;
		}
		distractors = uniqueDistractors;
	}



    // ------------------------------------------

    // DẠNG 6: TÍCH VÔ HƯỚNG & GÓC

    // ------------------------------------------

    else if (topic === 6 || topic === 'vector_angle') {
		distractors = [];
		const rand = Math.random();
		let subType = "";

		if (rand < 0.4) {
			subType = "special_angle";
			// =================================================================
			// LOẠI 1 (40%): BÀI TOÁN GÓC ĐẶC BIỆT (30°, 45°, 60°, 90°, 120°, 135°, 150°, 180°)
			// Yêu cầu: Tính số đo góc (độ)
			// =================================================================
			const specials = [
				{ deg: 30,  ax: "\\sqrt{3}", ay: "1",  bx: "1",  by: "\\sqrt{3}", dot: "2\\sqrt{3}", lenA: "2", lenB: "2", cos: "\\frac{\\sqrt{3}}{2}" },
				{ deg: 30,  ax: "\\sqrt{3}", ay: "1",  bx: "2",  by: "0",          dot: "2\\sqrt{3}", lenA: "2", lenB: "2", cos: "\\frac{\\sqrt{3}}{2}" },
				{ deg: 45,  ax: "1",          ay: "2",  bx: "-1", by: "3",          dot: "5",          lenA: "\\sqrt{5}", lenB: "\\sqrt{10}", cos: "\\frac{\\sqrt{2}}{2}" },
				{ deg: 45,  ax: "2",          ay: "1",  bx: "1",  by: "3",          dot: "5",          lenA: "\\sqrt{5}", lenB: "\\sqrt{10}", cos: "\\frac{\\sqrt{2}}{2}" },
				{ deg: 60,  ax: "1",          ay: "0",  bx: "1",  by: "\\sqrt{3}", dot: "1",          lenA: "1", lenB: "2", cos: "\\frac{1}{2}" },
				{ deg: 60,  ax: "\\sqrt{3}", ay: "1",  bx: "0",  by: "2",          dot: "2",          lenA: "2", lenB: "2", cos: "\\frac{1}{2}" },
				{ deg: 90,  ax: "2",          ay: "3",  bx: "-3", by: "2",          dot: "0",          lenA: "\\sqrt{13}", lenB: "\\sqrt{13}", cos: "0" },
				{ deg: 90,  ax: "1",          ay: "-2", bx: "2",  by: "1",          dot: "0",          lenA: "\\sqrt{5}", lenB: "\\sqrt{5}", cos: "0" },
				{ deg: 120, ax: "1",          ay: "0",  bx: "-1", by: "\\sqrt{3}", dot: "-1",         lenA: "1", lenB: "2", cos: "-\\frac{1}{2}" },
				{ deg: 120, ax: "\\sqrt{3}", ay: "1",  bx: "-\\sqrt{3}", by: "1", dot: "-2",         lenA: "2", lenB: "2", cos: "-\\frac{1}{2}" },
				{ deg: 135, ax: "1",          ay: "2",  bx: "1",  by: "-3",         dot: "-5",         lenA: "\\sqrt{5}", lenB: "\\sqrt{10}", cos: "-\\frac{\\sqrt{2}}{2}" },
				{ deg: 150, ax: "\\sqrt{3}", ay: "1",  bx: "-2", by: "0",          dot: "-2\\sqrt{3}", lenA: "2", lenB: "2", cos: "-\\frac{\\sqrt{3}}{2}" },
				{ deg: 150, ax: "\\sqrt{3}", ay: "1",  bx: "-1", by: "-\\sqrt{3}",dot: "-2\\sqrt{3}", lenA: "2", lenB: "2", cos: "-\\frac{\\sqrt{3}}{2}" },
				{ deg: 180, ax: "2",          ay: "-1", bx: "-4", by: "2",          dot: "-10",        lenA: "\\sqrt{5}", lenB: "\\sqrt{20}", cos: "-1" }
			];

			const item = specials[Math.floor(Math.random() * specials.length)];
			qText = `Góc giữa hai vectơ $\\vec{a} = (${item.ax}; ${item.ay})$ và $\\vec{b} = (${item.bx}; ${item.by})$ bằng:`;
			correctVal = `$${item.deg}^\\circ$`;

			// GÂY NHIỄU THEO BẪY TƯ DUY HỌC SINH:
			// Lỗi 1: Nhầm sang góc bù (180 - deg)
			const d1_deg = (item.deg !== 90) ? (180 - item.deg) : 45;
			distractors.push(`$${d1_deg}^\\circ$`);

			// Lỗi 2: Nhầm góc phụ (90 - deg) hoặc lệch góc 30 độ
			let d2_deg = Math.abs(90 - item.deg);
			if (d2_deg === 0 || d2_deg === item.deg || d2_deg === d1_deg) {
				d2_deg = (item.deg <= 120) ? item.deg + 30 : item.deg - 30;
			}
			distractors.push(`$${d2_deg}^\\circ$`);

			// Lỗi 3: Lấy góc đặc biệt khác từ danh sách bẫy [30, 45, 60, 90, 120, 135, 150, 180]
			const pool = [30, 45, 60, 90, 120, 135, 150, 180];
			for (let opt of pool) {
				let optStr = `$${opt}^\\circ$`;
				if (optStr !== correctVal && !distractors.includes(optStr)) {
					distractors.push(optStr);
					break;
				}
			}

			explanation = `Áp dụng công thức tính góc giữa hai vectơ:<br>` +
						  `$\\cos(\\vec{a}, \\vec{b}) = \\frac{\\vec{a}\\cdot\\vec{b}}{|\\vec{a}||\\vec{b}|} = \\frac{${item.dot}}{${item.lenA} \\cdot ${item.lenB}} = ${item.cos}$.<br>` +
						  `Do đó, góc giữa hai vectơ là $(\\vec{a}, \\vec{b}) = ${item.deg}^\\circ$.`;

		} else {
			subType = "general_cos";
			// =================================================================
			// LOẠI 2 (60%): BÀI TOÁN COSIN KHÔNG ĐẶC BIỆT
			// Yêu cầu: Tính cosin của góc giữa hai vectơ (Tự động rút gọn căn & phân thức KaTeX)
			// =================================================================
			let x1 = 0, y1 = 0, x2 = 0, y2 = 0;
			let dot = 0, sum1 = 0, sum2 = 0, prod = 0;

			while (true) {
				x1 = rNonZero(-4, 4);
				y1 = rNonZero(-4, 4);
				x2 = rNonZero(-4, 4);
				y2 = rNonZero(-4, 4);

				dot = x1 * x2 + y1 * y2;
				sum1 = x1 * x1 + y1 * y1;
				sum2 = x2 * x2 + y2 * y2;
				prod = sum1 * sum2;

				// Loại bỏ trường hợp vuông góc, cùng phương hoặc các góc đặc biệt
				if (dot === 0 || dot * dot === prod) continue;
				if (dot * dot * 4 === prod) continue;       // cos = +- 1/2 (60°, 120°)
				if (dot * dot * 2 === prod) continue;       // cos = +- sqrt(2)/2 (45°, 135°)
				if (dot * dot * 4 === prod * 3) continue;   // cos = +- sqrt(3)/2 (30°, 150°)

				break;
			}

			qText = `Cosin của góc giữa hai vectơ $\\vec{a} = (${x1}; ${y1})$ và $\\vec{b} = (${x2}; ${y2})$ bằng:`;
			correctVal = simplifyRadicalFraction(dot, sum1, sum2);

			// GÂY NHIỄU TRỰC TIẾP TỪ QUÁ TRÌNH GIẢI BÀI TOÁN:
			// Lỗi 1: Sai dấu tích vô hướng ở tử số (-dot)
			distractors.push(simplifyRadicalFraction(-dot, sum1, sum2));

			// Lỗi 2: Nhầm tích chéo ở tử số (x1*y2 + x2*y1)
			let wrongDot = x1 * y2 + x2 * y1;
			distractors.push(simplifyRadicalFraction(wrongDot, sum1, sum2));

			// Lỗi 3: Quên lấy căn bậc hai ở mẫu số (tính dot / (sum1 * sum2))
			distractors.push(simplifyFraction(dot, sum1 * sum2));

			explanation = `Công thức cosin góc giữa hai vectơ:<br>` +
						  `$\\cos(\\vec{a}, \\vec{b}) = \\frac{x_1 x_2 + y_1 y_2}{\\sqrt{x_1^2 + y_1^2} \\cdot \\sqrt{x_2^2 + y_2^2}} = \\frac{(${x1})(${x2}) + (${y1})(${y2})}{\\sqrt{(${x1})^2 + (${y1})^2} \\cdot \\sqrt{(${x2})^2 + (${y2})^2}} = \\frac{${dot}}{\\sqrt{${sum1}} \\cdot \\sqrt{${sum2}}} = $${correctVal}.`;
		}

		// Đảm bảo 3 phương án nhiễu duy nhất và khác đáp án đúng
		let uniqueDistractors = [];
		for (let d of distractors) {
			if (d !== correctVal && !uniqueDistractors.includes(d)) {
				uniqueDistractors.push(d);
			}
		}

		let backupVal = 1;
		while (uniqueDistractors.length < 3) {
			let candidate = (subType === "special_angle") 
				? `${(backupVal * 15) % 180}^$\\circ` 
				: simplifyFraction(backupVal, backupVal + 5);
			if (candidate !== correctVal && !uniqueDistractors.includes(candidate)) {
				uniqueDistractors.push(candidate);
			}
			backupVal++;
		}

		distractors = uniqueDistractors;
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

function checkAnswer(arg1, arg2) {
    // 1. Lấy đúng phần tử DOM theo ID thực tế trong file HTML (tương thích đa cấu hình)
    const optionsBox = document.getElementById('optionsBox') || document.getElementById('options-container');
    const feedbackBox = document.getElementById('feedbackBox') || document.getElementById('explanation') || document.getElementById('feedback-box');
    const nextBtn = document.getElementById('nextBtn');

    // Kiểm tra an toàn DOM
    if (!optionsBox) {
        console.error("Không tìm thấy phần tử chứa các nút đáp án (optionsBox)!");
        return;
    }

    // Lấy danh sách tất cả các nút phương án
    const buttons = optionsBox.querySelectorAll('button, .option-btn, .opt-btn');
    
    // Vô hiệu hóa các nút để tránh học sinh bấm nhiều lần trên màn hình cảm ứng điện thoại
    buttons.forEach(btn => btn.disabled = true);

    let selectedBtn = null;
    let selectedVal = "";

    // 2. Tương thích linh hoạt các kiểu tham số truyền vào từ sự kiện onclick
    if (arg1 instanceof HTMLElement) {
        selectedBtn = arg1;
        selectedVal = arg2 !== undefined ? arg2 : (arg1.getAttribute('data-val') || arg1.innerText);
    } else if (typeof arg1 === 'number') {
        selectedBtn = buttons[arg1];
        selectedVal = currentQuestion.options[arg1];
    } else {
        selectedVal = arg1;
        buttons.forEach(btn => {
            const dataVal = btn.getAttribute('data-val');
            if (dataVal === arg1 || btn.innerHTML.includes(arg1)) {
                selectedBtn = btn;
            }
        });
    }

    const cleanSelected = String(selectedVal).trim();
    const cleanCorrect = String(currentQuestion.correctVal || currentQuestion.correct).trim();
    const isCorrect = (cleanSelected === cleanCorrect);

    // 3. Hiệu ứng đổi màu nút đáp án trên giao diện điện thoại (Đúng: Xanh, Sai: Đỏ)
    buttons.forEach(btn => {
        const btnVal = btn.getAttribute('data-val') || btn.innerText;
        if (String(btnVal).trim() === cleanCorrect || btn.innerHTML.includes(cleanCorrect)) {
            btn.style.backgroundColor = "#2e7d32";
            btn.style.color = "#ffffff";
        }
    });

    if (!isCorrect && selectedBtn) {
        selectedBtn.style.backgroundColor = "#c62828";
        selectedBtn.style.color = "#ffffff";
    }

    // 4. Xây dựng giao diện Phản hồi & Lời giải chi tiết
    let statusHTML = "";
    if (isCorrect) {
        statusHTML = `
            <div style="background-color: #e8f5e9; border-left: 5px solid #2e7d32; padding: 12px; margin-top: 15px; border-radius: 8px;">
                <h4 style="color: #2e7d32; margin: 0 0 8px 0; font-size: 16px;">🎉 Chính xác! Bạn làm rất tốt.</h4>
                <p style="margin: 0; font-size: 14px; color: #1b5e20;"><b>💡 Lời giải chi tiết:</b> ${currentQuestion.explanation}</p>
            </div>
        `;
        if (typeof triggerHaptic === 'function') triggerHaptic('correct');
    } else {
        statusHTML = `
            <div style="background-color: #ffebee; border-left: 5px solid #c62828; padding: 12px; margin-top: 15px; border-radius: 8px;">
                <h4 style="color: #c62828; margin: 0 0 8px 0; font-size: 16px;">❌ Rất tiếc, câu trả lời chưa đúng!</h4>
                <div style="font-size: 14px; color: #b71c1c; margin-bottom: 6px;">
                    📌 <b>Bạn đã chọn:</b> ${selectedVal}
                </div>
                <div style="font-size: 14px; color: #1b5e20; margin-bottom: 8px;">
                    ✅ <b>Đáp án đúng là:</b> ${cleanCorrect}
                </div>
                <hr style="border: 0; border-top: 1px solid #ffcdd2; margin: 8px 0;">
                <p style="margin: 0; font-size: 14px; color: #333;"><b>💡 Lời giải chi tiết:</b> ${currentQuestion.explanation}</p>
            </div>
        `;
        if (typeof triggerHaptic === 'function') triggerHaptic('wrong');
    }

    // 5. Cập nhật DOM khung phản hồi & Kích hoạt Render lại KaTeX cho Phân số & Căn thức
    if (feedbackBox) {
        feedbackBox.innerHTML = statusHTML;
        feedbackBox.style.display = 'block';

        if (window.renderMathInElement) {
            renderMathInElement(feedbackBox, {
                delimiters: [
                    { left: "$$", right: "$$", display: true },
                    { left: "$", right: "$", display: false },
                    { left: "\\(", right: "\\)", display: false },
                    { left: "\\[", right: "\\]", display: true }
                ],
                throwOnError: false
            });
        }
    }

    // 6. Hiển thị nút chuyển câu tiếp theo ("#nextBtn")
    if (nextBtn) {
        nextBtn.style.display = "inline-block";
    }

    // 7. Cập nhật Điểm & Chuỗi đúng trong Chế độ Thử thách 20 câu
    if (typeof isChallenge !== 'undefined' && isChallenge) {
        if (isCorrect) {
            score += 1;
            streak += 1;
        } else {
            score = Math.max(0, score - 2);
            streak = 0;
        }
        if (typeof updateStats === 'function') updateStats();

        if (streak >= 20) {
            if (typeof triggerHaptic === 'function') triggerHaptic('win');
            setTimeout(() => {
                alert("🏆 CHÚC MỪNG! Bạn đã hoàn thành xuất sắc Thử thách 20 câu đúng liên tiếp!");
            }, 300);
        }
    }
}



function renderOptions(options, correctVal) {
    const container = document.getElementById('options-container');
    container.innerHTML = '';

    options.forEach((opt, idx) => {
        const btn = document.createElement('button');
        btn.className = 'opt-btn';
        btn.style.cssText = "width: 100%; margin-bottom: 10px; padding: 12px; font-size: 16px; border-radius: 8px; border: 1px solid #ccc; background-color: #fff; text-align: center;";
        btn.setAttribute('data-val', opt);
        btn.innerHTML = `<b>${String.fromCharCode(65 + idx)}.</b> ${opt}`;
        
        btn.onclick = () => checkAnswer(idx);
        container.appendChild(btn);
    });

    // Render KaTeX cho 4 nút đáp án trên giao diện điện thoại
    if (window.renderMathInElement) {
        renderMathInElement(container, {
            delimiters: [
                { left: "$$", right: "$$", display: true },
                { left: "$", right: "$", display: false }
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