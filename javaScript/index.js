let students = [];
let classroom = [];
let rows = 5;
let cols = 8;

// 初始化
document.addEventListener('DOMContentLoaded', function() {
    loadFromLocalStorage();
    updateClassroom();
    
    // 如果有排座数据，恢复排座显示
    if (classroom.length > 0) {
        updateClassroomDisplay();
    }
    
    updateStats();
    
    // 如果有缓存数据，显示提示
    const savedData = localStorage.getItem('classSeatingData');
    if (savedData) {
        try {
            const data = JSON.parse(savedData);
            if (data.students && data.students.length > 0) {
                showAlert(`已从缓存恢复 ${data.students.length} 名学生和排座数据`, 'info');
            }
        } catch (error) {
            console.error('解析缓存数据失败:', error);
        }
    }
});

// 保存数据到localStorage
function saveToLocalStorage() {
    try {
        const data = {
            students: students,
            classroom: classroom,
            rows: rows,
            cols: cols,
            timestamp: new Date().toISOString()
        };
        localStorage.setItem('classSeatingData', JSON.stringify(data));
        
        // 调试信息：显示保存的数据
        let seatCount = 0;
        for (let i = 0; i < classroom.length; i++) {
            for (let j = 0; j < classroom[i].length; j++) {
                if (classroom[i][j]) seatCount++;
            }
        }
        console.log(`数据已保存到localStorage: ${students.length}名学生, ${seatCount}个座位, ${rows}行×${cols}列`);
    } catch (error) {
        console.error('保存到localStorage失败:', error);
    }
}

// 从localStorage加载数据
function loadFromLocalStorage() {
    try {
        const savedData = localStorage.getItem('classSeatingData');
        if (savedData) {
            const data = JSON.parse(savedData);
            
            // 检查数据是否过期（7天）
            const savedTime = new Date(data.timestamp);
            const now = new Date();
            const daysDiff = (now - savedTime) / (1000 * 60 * 60 * 24);
            
            if (daysDiff > 7) {
                console.log('缓存数据已过期，清除旧数据');
                localStorage.removeItem('classSeatingData');
                return;
            }
            
            // 恢复数据
            if (data.students && Array.isArray(data.students)) {
                students = data.students;
                console.log(`从localStorage恢复了 ${students.length} 名学生`);
            }
            
            if (data.classroom && Array.isArray(data.classroom)) {
                classroom = data.classroom;
                
                // 统计恢复的排座数据
                let seatCount = 0;
                for (let i = 0; i < classroom.length; i++) {
                    for (let j = 0; j < classroom[i].length; j++) {
                        if (classroom[i][j]) seatCount++;
                    }
                }
                console.log(`从localStorage恢复了排座数据: ${seatCount}个座位`);
            }
            
            if (data.rows && data.cols) {
                rows = data.rows;
                cols = data.cols;
                
                // 更新输入框的值
                const rowsInput = document.getElementById('rows');
                const colsInput = document.getElementById('cols');
                if (rowsInput) rowsInput.value = rows;
                if (colsInput) colsInput.value = cols;
                
                console.log(`从localStorage恢复了教室布局: ${rows}行 × ${cols}列`);
            }
            
            // 如果有学生数据，更新学生列表显示
            if (students.length > 0) {
                updateStudentsList();
            }
            
            // 如果有排座数据，更新教室显示
            if (data.classroom && Array.isArray(data.classroom) && data.classroom.length > 0) {
                // 检查是否有实际的排座数据（不是空的）
                let hasSeatingData = false;
                for (let i = 0; i < data.classroom.length; i++) {
                    for (let j = 0; j < data.classroom[i].length; j++) {
                        if (data.classroom[i][j]) {
                            hasSeatingData = true;
                            break;
                        }
                    }
                    if (hasSeatingData) break;
                }
                
                if (hasSeatingData) {
                    console.log('检测到排座数据，将在updateClassroom后恢复');
                }
            }
        }
    } catch (error) {
        console.error('从localStorage加载数据失败:', error);
        // 如果加载失败，清除可能损坏的数据
        localStorage.removeItem('classSeatingData');
    }
}

// 显示提示信息
function showAlert(message, type = 'success') {
    const alert = document.getElementById('alert');
    alert.textContent = message;
    alert.className = `alert alert-${type}`;
    alert.style.display = 'block';
    
    // 使用 requestAnimationFrame 确保显示后再添加动画
    requestAnimationFrame(() => {
        alert.style.transform = 'translateY(0)';
    });
    
    setTimeout(() => {
        alert.style.transform = 'translateY(-100%)';
        setTimeout(() => {
            alert.style.display = 'none';
        }, 300);
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
    const newRows = parseInt(document.getElementById('rows').value);
    const newCols = parseInt(document.getElementById('cols').value);
    
    // 检查是否需要重新创建布局
    const needsRebuild = newRows !== rows || newCols !== cols;
    
    if (needsRebuild) {
        rows = newRows;
        cols = newCols;
        
        const classroomDiv = document.getElementById('classroom');
        classroomDiv.innerHTML = '';
        
        // 如果布局改变了，清空排座数据
        classroom = [];
        
        // 创建新的教室布局结构
        for (let i = 0; i < rows; i++) {
            classroom[i] = [];
            const rowDiv = document.createElement('div');
            rowDiv.className = 'classroom-row';
            
            // 创建座位容器
            const seatsContainer = document.createElement('div');
            seatsContainer.className = 'classroom-row-seats';
            // 动态设置列数，每2列座位为一组
            const seatGroupsCount = Math.ceil(cols / 2);
            seatsContainer.style.gridTemplateColumns = `repeat(${seatGroupsCount}, 1fr)`;
            
            for (let j = 0; j < cols; j += 2) {
                const seatGroup = document.createElement('div');
                seatGroup.className = 'seat-group';
                
                // 第一列座位
                const seat1 = document.createElement('div');
                seat1.className = 'seat empty';
                seat1.textContent = `${i + 1}-${j + 1}`;
                seat1.title = '点击安排学生';
                seat1.onclick = () => toggleSeat(i, j);
                seatGroup.appendChild(seat1);
                classroom[i][j] = null;
                
                // 第二列座位（同桌）
                if (j + 1 < cols) {
                    const seat2 = document.createElement('div');
                    seat2.className = 'seat empty';
                    seat2.textContent = `${i + 1}-${j + 2}`;
                    seat2.title = '点击安排学生';
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
        saveToLocalStorage(); // 保存到localStorage
    }
}

// 强制重建教室布局（用于手动更新布局按钮）
function rebuildClassroom() {
    const newRows = parseInt(document.getElementById('rows').value);
    const newCols = parseInt(document.getElementById('cols').value);
    
    rows = newRows;
    cols = newCols;
    
    const classroomDiv = document.getElementById('classroom');
    classroomDiv.innerHTML = '';
    
    // 清空排座数据
    classroom = [];
    
    // 创建新的教室布局结构
    for (let i = 0; i < rows; i++) {
        classroom[i] = [];
        const rowDiv = document.createElement('div');
        rowDiv.className = 'classroom-row';
        
        // 创建座位容器
        const seatsContainer = document.createElement('div');
        seatsContainer.className = 'classroom-row-seats';
        // 动态设置列数，每2列座位为一组
        const seatGroupsCount = Math.ceil(cols / 2);
        seatsContainer.style.gridTemplateColumns = `repeat(${seatGroupsCount}, 1fr)`;
        
        for (let j = 0; j < cols; j += 2) {
            const seatGroup = document.createElement('div');
            seatGroup.className = 'seat-group';
            
                            // 第一列座位
                const seat1 = document.createElement('div');
                seat1.className = 'seat empty';
                seat1.textContent = `${i + 1}-${j + 1}`;
                seat1.title = '点击安排学生';
                seat1.onclick = () => toggleSeat(i, j);
                seatGroup.appendChild(seat1);
                classroom[i][j] = null;
                
                // 第二列座位（同桌）
                if (j + 1 < cols) {
                    const seat2 = document.createElement('div');
                    seat2.className = 'seat empty';
                    seat2.textContent = `${i + 1}-${j + 2}`;
                    seat2.title = '点击安排学生';
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
    saveToLocalStorage(); // 保存到localStorage
}

// 清空特定座位
function clearSeat(row, col) {
    if (!classroom[row][col]) {
        showAlert('该座位本来就是空的', 'info');
        return;
    }
    
    const student = classroom[row][col];
    if (confirm(`确定要清空座位 ${row + 1}-${col + 1} 的学生 ${student.name} 吗？`)) {
        classroom[row][col] = null;
        updateSeatDisplay(row, col);
        saveToLocalStorage();
        showAlert(`已清空座位 ${row + 1}-${col + 1}`, 'success');
    }
}

// 安排学生到座位
function toggleSeat(row, col) {
    if (classroom[row][col]) {
        // 如果座位有人，显示提示信息
        const currentStudent = classroom[row][col];
        showAlert(`座位已被 ${currentStudent.name} 占用，无法修改。如需调整请先清空该学生`, 'info');
        return;
    } else {
        // 如果座位为空，可以选择学生
        if (students.length === 0) {
            showAlert('请先添加学生', 'error');
            return;
        }
        
        // 获取未安排的学生列表
        const unseatedStudents = getUnseatedStudents();
        if (unseatedStudents.length === 0) {
            showAlert('所有学生都已被安排到座位', 'info');
            return;
        }
        
        // 创建下拉选择框
        showStudentSelector(row, col, unseatedStudents);
    }
}

// 获取未安排到座位的学生列表
function getUnseatedStudents() {
    const seatedStudentNames = [];
    
    // 收集所有已安排到座位的学生姓名
    for (let i = 0; i < rows; i++) {
        for (let j = 0; j < cols; j++) {
            if (classroom[i][j]) {
                seatedStudentNames.push(classroom[i][j].name);
            }
        }
    }
    
    // 返回未安排到座位的学生
    return students.filter(student => !seatedStudentNames.includes(student.name));
}

// 显示学生选择器
function showStudentSelector(row, col, unseatedStudents) {
    // 创建模态框背景
    const modal = document.createElement('div');
    modal.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0, 0, 0, 0.5);
        z-index: 10000;
        display: flex;
        justify-content: center;
        align-items: center;
    `;
    
    // 创建选择框容器
    const selector = document.createElement('div');
    selector.style.cssText = `
        background: white;
        padding: 20px;
        border-radius: 8px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
        min-width: 300px;
        max-width: 400px;
    `;
    
    // 创建标题
    const title = document.createElement('h3');
    title.textContent = `选择要安排到座位 ${row + 1}-${col + 1} 的学生`;
    title.style.cssText = `
        margin: 0 0 15px 0;
        color: #333;
        font-size: 16px;
        text-align: center;
    `;
    
    // 创建下拉选择框
    const select = document.createElement('select');
    select.style.cssText = `
        width: 100%;
        padding: 10px;
        border: 1px solid #ddd;
        border-radius: 4px;
        font-size: 14px;
        margin-bottom: 15px;
    `;
    
    // 添加默认选项
    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.textContent = '请选择学生';
    select.appendChild(defaultOption);
    
    // 添加学生选项
    unseatedStudents.forEach(student => {
        const option = document.createElement('option');
        option.value = student.name;
        option.textContent = `${student.name} (${student.gender})`;
        select.appendChild(option);
    });
    
    // 创建按钮容器
    const buttonContainer = document.createElement('div');
    buttonContainer.style.cssText = `
        display: flex;
        gap: 10px;
        justify-content: center;
    `;
    
    // 创建确定按钮
    const confirmBtn = document.createElement('button');
    confirmBtn.textContent = '确定';
    confirmBtn.style.cssText = `
        padding: 8px 20px;
        background: #4facfe;
        color: white;
        border: none;
        border-radius: 4px;
        cursor: pointer;
        font-size: 14px;
    `;
    
    // 创建取消按钮
    const cancelBtn = document.createElement('button');
    cancelBtn.textContent = '取消';
    cancelBtn.style.cssText = `
        padding: 8px 20px;
        background: #6c757d;
        color: white;
        border: none;
        border-radius: 4px;
        cursor: pointer;
        font-size: 14px;
    `;
    
    // 添加按钮到容器
    buttonContainer.appendChild(confirmBtn);
    buttonContainer.appendChild(cancelBtn);
    
    // 组装选择器
    selector.appendChild(title);
    selector.appendChild(select);
    selector.appendChild(buttonContainer);
    
    // 添加到模态框
    modal.appendChild(selector);
    document.body.appendChild(modal);
    
    // 绑定事件
    confirmBtn.onclick = () => {
        const selectedStudentName = select.value;
        if (!selectedStudentName) {
            showAlert('请选择一个学生', 'error');
            return;
        }
        
        const student = students.find(s => s.name === selectedStudentName);
        if (student) {
            classroom[row][col] = student;
            updateSeatDisplay(row, col);
            saveToLocalStorage();
            showAlert(`学生 ${student.name} 已安排到该座位`, 'success');
        }
        
        // 关闭模态框
        document.body.removeChild(modal);
    };
    
    cancelBtn.onclick = () => {
        document.body.removeChild(modal);
    };
    
    // 点击背景关闭模态框
    modal.onclick = (e) => {
        if (e.target === modal) {
            document.body.removeChild(modal);
        }
    };
    
    // 按ESC键关闭模态框
    document.addEventListener('keydown', function closeOnEsc(e) {
        if (e.key === 'Escape') {
            if (document.body.contains(modal)) {
                document.body.removeChild(modal);
            }
            document.removeEventListener('keydown', closeOnEsc);
        }
    });
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
        saveToLocalStorage(); // 保存到localStorage
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
        saveToLocalStorage(); // 保存到localStorage
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
    saveToLocalStorage(); // 保存到localStorage
    
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
        saveToLocalStorage(); // 保存到localStorage
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
        saveToLocalStorage(); // 保存到localStorage
        showAlert('已清空所有学生数据', 'success');
    }
}

// 清除localStorage缓存
function clearLocalStorage() {
    if (confirm('确定要清除所有缓存数据吗？\n\n这将清除：\n• 学生名单\n• 排座结果\n• 教室布局设置\n\n此操作不可恢复！')) {
        try {
            localStorage.removeItem('classSeatingData');
            students = [];
            classroom = [];
            rows = 5;
            cols = 8;
            
            // 重置输入框
            const rowsInput = document.getElementById('rows');
            const colsInput = document.getElementById('cols');
            if (rowsInput) rowsInput.value = rows;
            if (colsInput) colsInput.value = cols;
            
            // 强制重建教室布局
            rebuildClassroom();
            updateStudentsList();
            updateStats();
            
            showAlert('已清除所有缓存数据，页面已重置', 'success');
        } catch (error) {
            console.error('清除localStorage失败:', error);
            showAlert('清除缓存失败，请手动刷新页面', 'error');
        }
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
    saveToLocalStorage(); // 保存到localStorage
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
        seat.title = `${classroom[row][col].name} (${classroom[row][col].gender}) - 右键清空座位`;
        
        // 移除之前的性别样式
        seat.classList.remove('male', 'female');
        
        // 添加性别样式
        if (classroom[row][col].gender === '男') {
            seat.classList.add('male');
        } else if (classroom[row][col].gender === '女') {
            seat.classList.add('female');
        }
        
        // 为已占用座位添加右键清空功能
        seat.oncontextmenu = (e) => {
            e.preventDefault();
            clearSeat(row, col);
        };
    } else {
        seat.className = 'seat empty';
        seat.textContent = `${row + 1}-${col + 1}`;
        seat.title = '点击安排学生';
        seat.classList.remove('male', 'female');
        
        // 移除右键事件
        seat.oncontextmenu = null;
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
        // 动态设置列数，每2列座位为一组
        const seatGroupsCount = Math.ceil(cols / 2);
        seatsContainer.style.gridTemplateColumns = `repeat(${seatGroupsCount}, 1fr)`;
        
        for (let j = 0; j < cols; j += 2) {
            const seatGroup = document.createElement('div');
            seatGroup.className = 'seat-group';
            
            // 第一列座位
            const seat1 = document.createElement('div');
            if (classroom[i][j]) {
                seat1.className = 'seat occupied';
                seat1.textContent = classroom[i][j].name;
                seat1.title = `${classroom[i][j].name} (${classroom[i][j].gender}) - 右键清空座位`;
                // 添加性别样式
                if (classroom[i][j].gender === '男') {
                    seat1.classList.add('male');
                } else if (classroom[i][j].gender === '女') {
                    seat1.classList.add('female');
                }
                // 为已占用座位添加右键清空功能
                seat1.oncontextmenu = (e) => {
                    e.preventDefault();
                    clearSeat(i, j);
                };
            } else {
                seat1.className = 'seat empty';
                seat1.textContent = `${i + 1}-${j + 1}`;
                seat1.title = '点击安排学生';
            }
            seat1.onclick = () => toggleSeat(i, j);
            seatGroup.appendChild(seat1);
            
            // 第二列座位（同桌）
            if (j + 1 < cols) {
                const seat2 = document.createElement('div');
                if (classroom[i][j + 1]) {
                    seat2.className = 'seat occupied';
                    seat2.textContent = classroom[i][j + 1].name;
                    seat2.title = `${classroom[i][j + 1].name} (${classroom[i][j + 1].gender}) - 右键清空座位`;
                    // 添加性别样式
                    if (classroom[i][j + 1].gender === '男') {
                        seat2.classList.add('male');
                    } else if (classroom[i][j + 1].gender === '女') {
                        seat2.classList.add('female');
                    }
                    // 为已占用座位添加右键清空功能
                    seat2.oncontextmenu = (e) => {
                        e.preventDefault();
                        clearSeat(i, j + 1);
                    };
                } else {
                    seat2.className = 'seat empty';
                    seat2.textContent = `${i + 1}-${j + 2}`;
                    seat2.title = '点击安排学生';
                }
                seat2.onclick = () => toggleSeat(i, j + 1);
                seatGroup.appendChild(seat2);
            }
            
            seatsContainer.appendChild(seatGroup);
        }
        
        rowDiv.appendChild(seatsContainer);
        classroomDiv.appendChild(rowDiv);
    }
    
    // 保存到localStorage
    saveToLocalStorage();
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
