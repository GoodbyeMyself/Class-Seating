let students = [];
let classroom = [];
let rows = 5;
let cols = 8;

// 初始化
document.addEventListener('DOMContentLoaded', function() {
    updateClassroom();
    updateStats();
});

// 显示提示信息
function showAlert(message, type = 'success') {
    const alert = document.getElementById('alert');
    alert.textContent = message;
    alert.className = `alert alert-${type}`;
    alert.style.display = 'block';
    
    setTimeout(() => {
        alert.style.display = 'none';
    }, 3000);
}

// 更新统计信息
function updateStats() {
    const total = students.length;
    const male = students.filter(s => s.gender === '男').length;
    const female = students.filter(s => s.gender === '女').length;
    
    document.getElementById('totalStudents').textContent = total;
    document.getElementById('maleStudents').textContent = male;
    document.getElementById('femaleStudents').textContent = female;
    document.getElementById('seatsCount').textContent = rows * cols;
}

// 更新教室布局
function updateClassroom() {
    rows = parseInt(document.getElementById('rows').value);
    cols = parseInt(document.getElementById('cols').value);
    
    const classroomDiv = document.getElementById('classroom');
    classroomDiv.innerHTML = '';
    classroom = [];
    
    // 创建新的教室布局结构
    for (let i = 0; i < rows; i++) {
        classroom[i] = [];
        const rowDiv = document.createElement('div');
        rowDiv.className = 'classroom-row';
        
        // 创建座位容器
        const seatsContainer = document.createElement('div');
        seatsContainer.className = 'classroom-row-seats';
        
        for (let j = 0; j < cols; j += 2) {
            const seatGroup = document.createElement('div');
            seatGroup.className = 'seat-group';
            
            // 第一列座位
            const seat1 = document.createElement('div');
            seat1.className = 'seat empty';
            seat1.textContent = `${i + 1}-${j + 1}`;
            seat1.onclick = () => toggleSeat(i, j);
            seatGroup.appendChild(seat1);
            classroom[i][j] = null;
            
            // 第二列座位（同桌）
            if (j + 1 < cols) {
                const seat2 = document.createElement('div');
                seat2.className = 'seat empty';
                seat2.textContent = `${i + 1}-${j + 2}`;
                seat2.onclick = () => toggleSeat(i, j + 1);
                seatGroup.appendChild(seat2);
                classroom[i][j + 1] = null;
            }
            
            seatsContainer.appendChild(seatGroup);
        }
        
        rowDiv.appendChild(seatsContainer);
        classroomDiv.appendChild(rowDiv);
    }
    
    updateStats();
}

// 切换座位状态
function toggleSeat(row, col) {
    if (classroom[row][col]) {
        // 如果座位有人，显示提示信息
        const currentStudent = classroom[row][col];
        showAlert(`座位已被 ${currentStudent.name} 占用`, 'info');
        return;
    } else {
        // 如果座位为空，可以选择学生
        if (students.length === 0) {
            showAlert('请先添加学生', 'error');
            return;
        }
        
        const studentName = prompt('请输入要安排到该座位的学生姓名：');
        if (studentName) {
            const student = students.find(s => s.name === studentName.trim());
            if (student) {
                // 检查学生是否已经在其他座位
                let isAlreadySeated = false;
                for (let i = 0; i < rows; i++) {
                    for (let j = 0; j < cols; j++) {
                        if (classroom[i][j] && classroom[i][j].name === student.name) {
                            isAlreadySeated = true;
                            break;
                        }
                    }
                }
                
                if (isAlreadySeated) {
                    showAlert(`学生 ${student.name} 已经在其他座位了`, 'error');
                } else {
                    classroom[row][col] = student;
                    updateSeatDisplay(row, col);
                    showAlert(`学生 ${student.name} 已安排到该座位`, 'success');
                }
            } else {
                showAlert('未找到该学生', 'error');
            }
        }
    }
}

// 从Excel导入数据
function importFromExcel() {
    const fileInput = document.getElementById('excelFile');
    const file = fileInput.files[0];
    
    if (!file) {
        showAlert('请先选择Excel文件', 'error');
        return;
    }
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
            
            // 检查是否是排座表格式（包含行/列标题）
            const isSeatingChart = jsonData.length > 0 && jsonData[0] && jsonData[0][0] === '行/列';
            
            if (isSeatingChart) {
                // 导入排座表格式
                showAlert('检测到排座表格式，正在导入座位安排...', 'info');
                importSeatingChart(jsonData);
            } else {
                // 导入学生名单格式
                showAlert('检测到学生名单格式，正在导入学生信息...', 'info');
                importStudentList(jsonData);
            }
            
            fileInput.value = '';
        } catch (error) {
            showAlert('Excel文件读取失败，请检查文件格式', 'error');
            console.error(error);
        }
    };
    
    reader.readAsArrayBuffer(file);
}

// 导入学生名单格式
function importStudentList(jsonData) {
    const newStudents = [];
    for (let i = 1; i < jsonData.length; i++) { // 跳过标题行
        const row = jsonData[i];
        if (row[0] && row[1]) {
            newStudents.push({
                name: row[0].toString().trim(),
                gender: row[1].toString().trim()
            });
        }
    }
    
    if (newStudents.length > 0) {
        // 检查是否已存在学生列表
        if (students.length > 0) {
            const choice = confirm(`当前已有 ${students.length} 名学生，是否要清空现有学生列表并导入新的 ${newStudents.length} 名学生？\n\n点击"确定"清空现有列表并导入新学生\n点击"取消"将新学生添加到现有列表中`);
            
            if (choice) {
                // 清空现有学生列表和教室
                students = [];
                for (let i = 0; i < rows; i++) {
                    for (let j = 0; j < cols; j++) {
                        classroom[i][j] = null;
                    }
                }
                students = [...newStudents];
                updateClassroomDisplay();
                showAlert(`已清空现有学生列表，成功导入 ${newStudents.length} 名新学生`, 'success');
            } else {
                // 添加到现有列表
                students = [...students, ...newStudents];
                showAlert(`已将 ${newStudents.length} 名新学生添加到现有列表中`, 'success');
            }
        } else {
            // 没有现有学生，直接导入
            students = [...newStudents];
            showAlert(`成功导入 ${newStudents.length} 名学生`, 'success');
        }
        
        updateStudentsList();
        updateStats();
    } else {
        showAlert('Excel文件中没有找到有效数据', 'error');
    }
}

// 导入排座表格式
function importSeatingChart(jsonData) {
    // 检查是否已存在学生列表
    if (students.length > 0) {
        const choice = confirm(`当前已有 ${students.length} 名学生，导入排座表将清空现有学生列表和座位安排。\n\n是否继续导入？\n\n点击"确定"清空现有数据并导入排座表\n点击"取消"取消导入操作`);
        
        if (!choice) {
            showAlert('已取消导入排座表', 'info');
            return;
        }
    }
    
    // 清空当前教室和学生列表
    students = [];
    for (let i = 0; i < rows; i++) {
        for (let j = 0; j < cols; j++) {
            classroom[i][j] = null;
        }
    }
    
    // 从排座表数据中提取学生信息并安排座位
    let importedStudents = [];
    let seatCount = 0;
    
    // 跳过标题行，从第2行开始读取座位数据
    for (let i = 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (row && row[0] && row[0].includes('行')) {
            // 这是座位行
            const rowIndex = i - 1; // 实际行索引
            
            for (let j = 1; j < row.length && j <= cols; j++) {
                const cellValue = row[j];
                if (cellValue && cellValue !== '空座位' && !cellValue.includes('行') && !cellValue.includes('列')) {
                    // 提取学生姓名和性别
                    const studentInfo = parseStudentInfo(cellValue);
                    if (studentInfo) {
                        // 检查学生是否已存在，如果不存在则添加到学生列表
                        let existingStudent = students.find(s => s.name === studentInfo.name);
                        if (!existingStudent) {
                            students.push(studentInfo);
                            importedStudents.push(studentInfo);
                        } else {
                            // 使用已存在的学生信息
                            studentInfo.gender = existingStudent.gender;
                        }
                        
                        // 安排到座位
                        const colIndex = j - 1; // 实际列索引
                        if (rowIndex < rows && colIndex < cols) {
                            classroom[rowIndex][colIndex] = studentInfo;
                            seatCount++;
                        }
                    }
                }
            }
        }
    }
    
    if (seatCount > 0) {
        // 更新显示
        updateStudentsList();
        updateStats();
        updateClassroomDisplay();
        
        if (importedStudents.length > 0) {
            showAlert(`成功导入排座表，安排了 ${seatCount} 个座位，新增 ${importedStudents.length} 名学生`, 'success');
        } else {
            showAlert(`成功导入排座表，安排了 ${seatCount} 个座位`, 'success');
        }
    } else {
        showAlert('排座表中没有找到有效的座位数据', 'error');
    }
}

// 解析学生信息
function parseStudentInfo(cellValue) {
    if (!cellValue || typeof cellValue !== 'string') return null;
    
    // 处理换行符分隔的格式：姓名\n(性别)
    const lines = cellValue.split('\n');
    if (lines.length >= 2) {
        const name = lines[0].trim();
        const genderMatch = lines[1].match(/\(([男女])\)/);
        if (name && genderMatch) {
            return {
                name: name,
                gender: genderMatch[1]
            };
        }
    }
    
    // 如果没有换行符，尝试其他格式
    const genderMatch = cellValue.match(/\(([男女])\)/);
    if (genderMatch) {
        const name = cellValue.replace(/\([男女]\)/, '').trim();
        if (name) {
            return {
                name: name,
                gender: genderMatch[1]
            };
        }
    }
    
    // 如果只包含姓名，默认设置为男性（可以根据需要调整）
    if (cellValue.trim()) {
        return {
            name: cellValue.trim(),
            gender: '男' // 默认性别
        };
    }
    
    return null;
}

// 手动添加学生
function addStudent() {
    const name = document.getElementById('studentName').value.trim();
    const gender = document.getElementById('studentGender').value;
    
    if (!name || !gender) {
        showAlert('请填写完整的学生信息', 'error');
        return;
    }
    
    if (students.some(s => s.name === name)) {
        showAlert('该学生已存在', 'error');
        return;
    }
    
    students.push({ name, gender });
    updateStudentsList();
    updateStats();
    
    // 清空输入框
    document.getElementById('studentName').value = '';
    document.getElementById('studentGender').value = '';
    
    showAlert(`成功添加学生：${name}`, 'success');
}

// 更新学生列表显示
function updateStudentsList() {
    const listDiv = document.getElementById('studentsList');
    
    if (students.length === 0) {
        listDiv.innerHTML = '<p style="text-align: center; color: #999; padding: 15px; font-size: 13px;">暂无学生数据</p>';
        return;
    }
    
    listDiv.innerHTML = students.map((student, index) => `
        <div class="student-item">
            <div class="student-info">
                <span style="font-weight: bold; font-size: 13px;">${student.name}</span>
                <span class="gender-badge gender-${student.gender === '男' ? 'male' : 'female'}">${student.gender}</span>
            </div>
            <button class="btn btn-danger" onclick="removeStudent(${index})" style="padding: 3px 8px; font-size: 11px;">删除</button>
        </div>
    `).join('');
}

// 删除学生
function removeStudent(index) {
    const student = students[index];
    if (confirm(`确定要删除学生 ${student.name} 吗？`)) {
        students.splice(index, 1);
        updateStudentsList();
        updateStats();
        
        // 从教室中移除该学生
        for (let i = 0; i < rows; i++) {
            for (let j = 0; j < cols; j++) {
                if (classroom[i][j] && classroom[i][j].name === student.name) {
                    classroom[i][j] = null;
                    updateSeatDisplay(i, j);
                }
            }
        }
        showAlert(`已删除学生：${student.name}`, 'success');
    }
}

// 清空所有学生
function clearAllStudents() {
    if (students.length === 0) {
        showAlert('当前没有学生数据', 'error');
        return;
    }
    
    if (confirm('确定要清空所有学生数据吗？此操作不可恢复！')) {
        students = [];
        updateStudentsList();
        updateStats();
        
        // 清空教室
        for (let i = 0; i < rows; i++) {
            for (let j = 0; j < cols; j++) {
                classroom[i][j] = null;
                updateSeatDisplay(i, j);
            }
        }
        showAlert('已清空所有学生数据', 'success');
    }
}

// 自动排座
function autoArrange() {
    if (students.length === 0) {
        showAlert('请先添加学生', 'error');
        return;
    }
    
    if (students.length > rows * cols) {
        showAlert(`学生数量(${students.length})超过座位数量(${rows * cols})`, 'error');
        return;
    }
    
    // 清空教室
    for (let i = 0; i < rows; i++) {
        for (let j = 0; j < cols; j++) {
            classroom[i][j] = null;
        }
    }
    
    // 随机打乱学生顺序
    const shuffledStudents = [...students].sort(() => Math.random() - 0.5);
    
    // 按行优先安排座位
    let studentIndex = 0;
    for (let i = 0; i < rows && studentIndex < shuffledStudents.length; i++) {
        for (let j = 0; j < cols && studentIndex < shuffledStudents.length; j++) {
            classroom[i][j] = shuffledStudents[studentIndex];
            studentIndex++;
        }
    }
    
    // 更新显示
    updateClassroomDisplay();
    showAlert('自动排座完成！', 'success');
}

// 更新特定座位的显示
function updateSeatDisplay(row, col) {
    const classroomDiv = document.getElementById('classroom');
    if (!classroomDiv) return;
    
    const rowDiv = classroomDiv.children[row];
    if (!rowDiv) return;
    
    const seatsContainer = rowDiv.querySelector('.classroom-row-seats');
    if (!seatsContainer) return;
    
    // 计算座位在座位组中的位置
    const seatGroupIndex = Math.floor(col / 2);
    const seatGroup = seatsContainer.children[seatGroupIndex];
    if (!seatGroup) return;
    
    // 确定是座位组中的第一个还是第二个座位
    const isFirstSeat = col % 2 === 0;
    const seat = isFirstSeat ? seatGroup.children[0] : seatGroup.children[1];
    
    if (!seat) return;
    
    if (classroom[row][col]) {
        seat.className = 'seat occupied';
        seat.textContent = classroom[row][col].name;
        seat.title = `${classroom[row][col].name} (${classroom[row][col].gender})`;
        
        // 移除之前的性别样式
        seat.classList.remove('male', 'female');
        
        // 添加性别样式
        if (classroom[row][col].gender === '男') {
            seat.classList.add('male');
        } else if (classroom[row][col].gender === '女') {
            seat.classList.add('female');
        }
    } else {
        seat.className = 'seat empty';
        seat.textContent = `${row + 1}-${col + 1}`;
        seat.title = '';
        seat.classList.remove('male', 'female');
    }
    
    // 重新绑定点击事件
    seat.onclick = () => toggleSeat(row, col);
}

// 更新教室显示
function updateClassroomDisplay() {
    const classroomDiv = document.getElementById('classroom');
    classroomDiv.innerHTML = '';
    
    for (let i = 0; i < rows; i++) {
        const rowDiv = document.createElement('div');
        rowDiv.className = 'classroom-row';
        
        // 创建座位容器
        const seatsContainer = document.createElement('div');
        seatsContainer.className = 'classroom-row-seats';
        
        for (let j = 0; j < cols; j += 2) {
            const seatGroup = document.createElement('div');
            seatGroup.className = 'seat-group';
            
            // 第一列座位
            const seat1 = document.createElement('div');
            if (classroom[i][j]) {
                seat1.className = 'seat occupied';
                seat1.textContent = classroom[i][j].name;
                seat1.title = `${classroom[i][j].name} (${classroom[i][j].gender})`;
                // 添加性别样式
                if (classroom[i][j].gender === '男') {
                    seat1.classList.add('male');
                } else if (classroom[i][j].gender === '女') {
                    seat1.classList.add('female');
                }
            } else {
                seat1.className = 'seat empty';
                seat1.textContent = `${i + 1}-${j + 1}`;
            }
            seat1.onclick = () => toggleSeat(i, j);
            seatGroup.appendChild(seat1);
            
            // 第二列座位（同桌）
            if (j + 1 < cols) {
                const seat2 = document.createElement('div');
                if (classroom[i][j + 1]) {
                    seat2.className = 'seat occupied';
                    seat2.textContent = classroom[i][j + 1].name;
                    seat2.title = `${classroom[i][j + 1].name} (${classroom[i][j + 1].gender})`;
                    // 添加性别样式
                    if (classroom[i][j + 1].gender === '男') {
                        seat2.classList.add('male');
                    } else if (classroom[i][j + 1].gender === '女') {
                        seat2.classList.add('female');
                    }
                } else {
                    seat2.className = 'seat empty';
                    seat2.textContent = `${i + 1}-${j + 2}`;
                }
                seat2.onclick = () => toggleSeat(i, j + 1);
                seatGroup.appendChild(seat2);
            }
            
            seatsContainer.appendChild(seatGroup);
        }
        
        rowDiv.appendChild(seatsContainer);
        classroomDiv.appendChild(rowDiv);
    }
}

// 导出排座图
function exportSeatingChart() {
    if (students.length === 0) {
        showAlert('请先添加学生', 'error');
        return;
    }
    
    // 准备Excel数据 - 按照教室座位布局生成
    let excelData = [];
    
    // 添加列标题行
    let headerRow = ['行/列'];
    for (let j = 0; j < cols; j++) {
        headerRow.push(`第${j + 1}列`);
    }
    excelData.push(headerRow);
    
    // 添加每行的数据
    for (let i = 0; i < rows; i++) {
        let rowData = [`第${i + 1}行`];
        for (let j = 0; j < cols; j++) {
            if (classroom[i][j]) {
                // 有学生的座位，显示姓名和性别
                rowData.push(`${classroom[i][j].name}\n(${classroom[i][j].gender})`);
            } else {
                // 空座位
                rowData.push('空座位');
            }
        }
        excelData.push(rowData);
    }
    
    // 添加统计信息行
    excelData.push([]); // 空行
    excelData.push(['统计信息']);
    excelData.push(['总学生数', students.length]);
    excelData.push(['男生人数', students.filter(s => s.gender === '男').length]);
    excelData.push(['女生人数', students.filter(s => s.gender === '女').length]);
    excelData.push(['总座位数', rows * cols]);
    excelData.push(['空闲座位数', rows * cols - students.length]);
    
    // 创建工作簿
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(excelData);
    
    // 设置列宽 - 根据列数动态设置
    let colWidths = [{ wch: 12 }]; // 行/列标题列
    for (let j = 0; j < cols; j++) {
        colWidths.push({ wch: 15 }); // 每列座位
    }
    ws['!cols'] = colWidths;
    
    // 设置行高
    let rowHeights = [];
    for (let i = 0; i < rows + 8; i++) { // +8 包括标题行和统计信息
        rowHeights.push({ hpt: 25 }); // 设置行高为25磅
    }
    ws['!rows'] = rowHeights;
    
    // 添加工作表到工作簿
    XLSX.utils.book_append_sheet(wb, ws, '班级排座表');
    
    // 导出Excel文件
    XLSX.writeFile(wb, '班级排座表.xlsx');
    
    showAlert('排座表已导出为Excel文件', 'success');
}

// 打印排座图
function printSeatingChart() {
    if (students.length === 0) {
        showAlert('请先添加学生', 'error');
        return;
    }
    
    const printWindow = window.open('', '_blank');
    printWindow.document.write(`
        <html>
        <head>
            <title>班级排座表</title>
            <style>
                body { font-family: Arial, sans-serif; margin: 20px; }
                .header { text-align: center; margin-bottom: 30px; }
                .classroom { display: grid; grid-template-columns: repeat(${cols}, 1fr); gap: 10px; margin: 20px 0; }
                .seat { border: 1px solid #000; padding: 10px; text-align: center; min-height: 40px; }
                .occupied { background: #e3f2fd; }
                .empty { background: #f5f5f5; color: #999; }
                .legend { margin-top: 20px; }
                .legend-item { display: inline-block; margin-right: 20px; }
            </style>
        </head>
        <body>
            <div class="header">
                <h1>班级排座表</h1>
                <p>生成时间：${new Date().toLocaleString()}</p>
            </div>
            <div class="classroom">
    `);
    
    for (let i = 0; i < rows; i++) {
        for (let j = 0; j < cols; j++) {
            if (classroom[i][j]) {
                printWindow.document.write(`
                    <div class="seat occupied">
                        <strong>${classroom[i][j].name}</strong><br>
                        <small>${classroom[i][j].gender}</small>
                    </div>
                `);
            } else {
                printWindow.document.write(`
                    <div class="seat empty">座位 ${i + 1}-${j + 1}</div>
                `);
            }
        }
    }
    
    printWindow.document.write(`
            </div>
            <div class="legend">
                <div class="legend-item">总学生数：${students.length}</div>
                <div class="legend-item">男生：${students.filter(s => s.gender === '男').length}</div>
                <div class="legend-item">女生：${students.filter(s => s.gender === '女').length}</div>
            </div>
        </body>
        </html>
    `);
    
    printWindow.document.close();
    printWindow.print();
}

// 文件拖拽功能
document.addEventListener('DOMContentLoaded', function() {
    const fileInput = document.querySelector('.file-input');
    
    fileInput.addEventListener('dragover', (e) => {
        e.preventDefault();
        fileInput.style.borderColor = '#4facfe';
        fileInput.style.background = '#f0f8ff';
    });
    
    fileInput.addEventListener('dragleave', (e) => {
        e.preventDefault();
        fileInput.style.borderColor = '#ddd';
        fileInput.style.background = '#f8f9fa';
    });
    
    fileInput.addEventListener('drop', (e) => {
        e.preventDefault();
        fileInput.style.borderColor = '#ddd';
        fileInput.style.background = '#f8f9fa';
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            document.getElementById('excelFile').files = files;
            // 拖拽后自动导入
            importFromExcel();
        }
    });
});
