// 全局变量
let currentRoomType = 4; // 默认四人寝
let currentGender = 'male'; // 默认男寝

// 为每个寝室类型和性别独立存储数据和状态
let roomTypeData = {
    2: {
        male: {
            excelFile: null,
            imageFile: null,
            currentImages: [],
            currentPage: 0,
            totalStudents: 0,
            studentsPerCard: 0,
            hasResults: false,
            dormInfo: null,
            collegeName: null
        },
        female: {
            excelFile: null,
            imageFile: null,
            currentImages: [],
            currentPage: 0,
            totalStudents: 0,
            studentsPerCard: 0,
            hasResults: false,
            dormInfo: null,
            collegeName: null
        }
    },
    3: {
        male: {
            excelFile: null,
            imageFile: null,
            currentImages: [],
            currentPage: 0,
            totalStudents: 0,
            studentsPerCard: 0,
            hasResults: false,
            dormInfo: null,
            collegeName: null
        },
        female: {
            excelFile: null,
            imageFile: null,
            currentImages: [],
            currentPage: 0,
            totalStudents: 0,
            studentsPerCard: 0,
            hasResults: false,
            dormInfo: null,
            collegeName: null
        }
    },
    4: {
        male: {
            excelFile: null,
            imageFile: null,
            currentImages: [],
            currentPage: 0,
            totalStudents: 0,
            studentsPerCard: 0,
            hasResults: false,
            dormInfo: null,
            collegeName: null
        },
        female: {
            excelFile: null,
            imageFile: null,
            currentImages: [],
            currentPage: 0,
            totalStudents: 0,
            studentsPerCard: 0,
            hasResults: false,
            dormInfo: null,
            collegeName: null
        }
    },
    5: {
        male: {
            excelFile: null,
            imageFile: null,
            currentImages: [],
            currentPage: 0,
            totalStudents: 0,
            studentsPerCard: 0,
            hasResults: false,
            dormInfo: null,
            collegeName: null
        },
        female: {
            excelFile: null,
            imageFile: null,
            currentImages: [],
            currentPage: 0,
            totalStudents: 0,
            studentsPerCard: 0,
            hasResults: false,
            dormInfo: null,
            collegeName: null
        }
    }
};
let positionAdjustments = {
    2: {
        1: { photo_x: 0, photo_y: 0, photo_width: 0, photo_height: 0, text_x: 0, text_y: 0, font_size: 0, line_spacing: 0 },
        2: { photo_x: 0, photo_y: 0, photo_width: 0, photo_height: 0, text_x: 0, text_y: 0, font_size: 0, line_spacing: 0 }
    },
    3: {
        1: { photo_x: 0, photo_y: 0, photo_width: 0, photo_height: 0, text_x: 0, text_y: 0, font_size: 0, line_spacing: 0 },
        2: { photo_x: 0, photo_y: 0, photo_width: 0, photo_height: 0, text_x: 0, text_y: 0, font_size: 0, line_spacing: 0 },
        3: { photo_x: 0, photo_y: 0, photo_width: 0, photo_height: 0, text_x: 0, text_y: 0, font_size: 0, line_spacing: 0 }
    },
    4: {
        1: { photo_x: 0, photo_y: 0, photo_width: 0, photo_height: 0, text_x: 0, text_y: 0, font_size: 0, line_spacing: 0 },
        2: { photo_x: 0, photo_y: 0, photo_width: 0, photo_height: 0, text_x: 0, text_y: 0, font_size: 0, line_spacing: 0 },
        3: { photo_x: 0, photo_y: 0, photo_width: 0, photo_height: 0, text_x: 0, text_y: 0, font_size: 0, line_spacing: 0 },
        4: { photo_x: 0, photo_y: 0, photo_width: 0, photo_height: 0, text_x: 0, text_y: 0, font_size: 0, line_spacing: 0 }
    },
    5: {
        1: { photo_x: 0, photo_y: 0, photo_width: 0, photo_height: 0, text_x: 0, text_y: 0, font_size: 0, line_spacing: 0 },
        2: { photo_x: 0, photo_y: 0, photo_width: 0, photo_height: 0, text_x: 0, text_y: 0, font_size: 0, line_spacing: 0 },
        3: { photo_x: 0, photo_y: 0, photo_width: 0, photo_height: 0, text_x: 0, text_y: 0, font_size: 0, line_spacing: 0 },
        4: { photo_x: 0, photo_y: 0, photo_width: 0, photo_height: 0, text_x: 0, text_y: 0, font_size: 0, line_spacing: 0 },
        5: { photo_x: 0, photo_y: 0, photo_width: 0, photo_height: 0, text_x: 0, text_y: 0, font_size: 0, line_spacing: 0 }
    }
};

// DOM 元素引用
const excelUpload = document.getElementById('excelUpload');
const imageUpload = document.getElementById('imageUpload');
const excelFileInput = document.getElementById('excelFile');
const imageFileInput = document.getElementById('imageFile');
const excelInfo = document.getElementById('excelInfo');
const imageInfo = document.getElementById('imageInfo');
const generateBtn = document.getElementById('generateBtn');
const resultSection = document.getElementById('resultSection');
const resultInfo = document.getElementById('resultInfo');
const paginationControls = document.getElementById('paginationControls');
const prevBtn = document.getElementById('prevBtn');
const nextBtn = document.getElementById('nextBtn');
const pageSelect = document.getElementById('pageSelect');
const totalPages = document.getElementById('totalPages');
const currentCard = document.getElementById('currentCard');
const downloadBtn = document.getElementById('downloadBtn');

// 位置配置区域
const positionConfigSection = document.getElementById('positionConfigSection');

// 操作按钮
const applyAllBtn = document.getElementById('applyAllBtn');
const resetAllBtn = document.getElementById('resetAllBtn');

// 动态获取输入框元素的函数
function getInputElement(positionNum, paramName, roomType) {
    const suffix = `_${roomType}`;
    const elementId = `position${positionNum}${paramName}${suffix}`;
    return document.getElementById(elementId);
}

// 获取当前寝室类型和性别的数据
function getCurrentData() {
    return roomTypeData[currentRoomType][currentGender];
}

// 初始化事件监听器
document.addEventListener('DOMContentLoaded', function () {
    initializeUploadAreas();
    initializeButtons();
    initializePagination();
    initializePositionAdjustments();
    initializeNavigation();
    // 初始化默认显示四人寝的位置调整区域
    switchRoomType(4);
    // 初始化默认性别为男寝
    switchGender('male');
    // 初始化默认位置参数
    initializeDefaultPositions();
});

// 初始化上传区域
function initializeUploadAreas() {
    // Excel上传区域
    setupUploadArea(excelUpload, excelFileInput, handleExcelFile, 'excel');
    
    // 图片上传区域
    setupUploadArea(imageUpload, imageFileInput, handleImageFile, 'image');
}

// 设置上传区域的拖拽功能
function setupUploadArea(uploadArea, fileInput, handleFile, type) {
    // 点击上传
    uploadArea.addEventListener('click', () => {
        fileInput.click();
    });

    // 文件选择
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
        }
    });

    // 拖拽事件
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('drag-over');
    });

    uploadArea.addEventListener('dragleave', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('drag-over');
    });

    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('drag-over');
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleFile(files[0]);
        }
    });
}

// 处理Excel文件
function handleExcelFile(file) {
    if (!isValidExcelFile(file)) {
        showError('请上传有效的Excel文件 (.xlsx, .xls)');
        return;
    }
    
    getCurrentData().excelFile = file;
    excelInfo.textContent = `已选择: ${file.name} (${formatFileSize(file.size)})`;
    excelUpload.classList.add('file-selected');
    
    checkFilesReady();
}

// 处理图片文件
function handleImageFile(file) {
    if (!isValidImageFile(file)) {
        showError('请上传有效的图片文件 (.jpg, .jpeg, .png)');
        return;
    }
    
    getCurrentData().imageFile = file;
    imageInfo.textContent = `已选择: ${file.name} (${formatFileSize(file.size)})`;
    imageUpload.classList.add('file-selected');
    
    checkFilesReady();
}

// 验证Excel文件
function isValidExcelFile(file) {
    const validTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                       'application/vnd.ms-excel'];
    const validExtensions = ['.xlsx', '.xls'];
    
    return validTypes.includes(file.type) || 
           validExtensions.some(ext => file.name.toLowerCase().endsWith(ext));
}

// 验证图片文件
function isValidImageFile(file) {
    const validTypes = ['image/jpeg', 'image/jpg', 'image/png'];
    return validTypes.includes(file.type);
}

// 格式化文件大小
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// 检查文件是否准备就绪
function checkFilesReady() {
    const currentData = getCurrentData();
    if (currentData.excelFile && currentData.imageFile) {
        generateBtn.disabled = false;
        generateBtn.classList.add('ready');
    } else {
        generateBtn.disabled = true;
        generateBtn.classList.remove('ready');
    }
}

// 初始化按钮事件
function initializeButtons() {
    generateBtn.addEventListener('click', generateCards);
    downloadBtn.addEventListener('click', downloadAllImages);
    
    // 添加ZIP下载按钮事件监听器
    const downloadZipBtn = document.getElementById('downloadZipBtn');
    if (downloadZipBtn) {
        downloadZipBtn.addEventListener('click', downloadZipFile);
    }
}

// 初始化翻页功能
function initializePagination() {
    prevBtn.addEventListener('click', () => {
        const currentData = getCurrentData();
        if (currentData.currentPage > 0) {
            currentData.currentPage--;
            displayCurrentCard();
            updatePaginationControls();
        }
    });

    nextBtn.addEventListener('click', () => {
        const currentData = getCurrentData();
        if (currentData.currentPage < currentData.currentImages.length - 1) {
            currentData.currentPage++;
            displayCurrentCard();
            updatePaginationControls();
        }
    });

    // 添加页面选择器事件监听
    pageSelect.addEventListener('change', (e) => {
        const currentData = getCurrentData();
        const selectedPage = parseInt(e.target.value) - 1;
        if (selectedPage >= 0 && selectedPage < currentData.currentImages.length) {
            currentData.currentPage = selectedPage;
            displayCurrentCard();
            updatePaginationControls();
        }
    });
}

// 生成信息卡
async function generateCards() {
    const currentData = getCurrentData();
    if (!currentData.excelFile || !currentData.imageFile) {
        showError('请先上传Excel文件和背景图片');
        return;
    }

    // 显示加载状态
    setLoadingState(true);

    try {
        const formData = new FormData();
        formData.append('excel_file', currentData.excelFile);
        formData.append('image_file', currentData.imageFile);
        formData.append('room_type', currentRoomType); // 添加寝室类型参数
        
        // 添加当前寝室类型的位置调整参数
        if (positionAdjustments[currentRoomType]) {
            formData.append('position_adjustments', JSON.stringify(positionAdjustments[currentRoomType]));
        }

        const response = await fetch('/generate', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const result = await response.json();
        
        if (result.success) {
            // 保存当前页面状态
            const previousPage = currentData.currentPage;
            
            // 保存结果到当前寝室类型的数据中
            currentData.currentImages = result.images;
            currentData.totalStudents = result.total_students;
            currentData.studentsPerCard = result.students_per_card;
            currentData.dormInfo = result.dorm_info; // 保存寝室信息
            currentData.collegeName = result.college_name; // 保存学院信息
            currentData.hasResults = true;
            
            // 如果之前有结果且当前页面在有效范围内，保持当前页面；否则重置为第一页
            if (previousPage < result.images.length) {
                currentData.currentPage = previousPage;
            } else {
                currentData.currentPage = 0;
            }
            
            displayResults(result.images, result.total_students, result.students_per_card);
            showSuccess(`成功生成 ${result.images.length} 张信息卡，共 ${result.total_students} 名学生`);
        } else {
            throw new Error(result.error || '生成失败');
        }
    } catch (error) {
        console.error('生成信息卡时出错:', error);
        showError('生成信息卡时出错: ' + error.message);
    } finally {
        setLoadingState(false);
    }
}

// 设置加载状态
function setLoadingState(loading) {
    const btnText = generateBtn.querySelector('.btn-text');
    const spinner = generateBtn.querySelector('.loading-spinner');
    
    if (loading) {
        btnText.textContent = '生成中...';
        spinner.style.display = 'inline-block';
        generateBtn.disabled = true;
    } else {
        btnText.textContent = '生成信息卡';
        spinner.style.display = 'none';
        generateBtn.disabled = false;
    }
}

// 显示结果
function displayResults(images, totalStudents, studentsPerCard) {
    const currentData = getCurrentData();
    currentData.currentImages = images;
    
    // 保持当前页面位置，但确保不超出新的页面范围
    if (currentData.currentPage >= images.length) {
        currentData.currentPage = Math.max(0, images.length - 1);
    }
    
    // 显示结果信息
    resultInfo.innerHTML = `
        <div class="result-summary">
            <span class="result-count">成功生成 ${images.length} 张学生信息卡</span>
            <span class="student-info">共 ${totalStudents} 名学生，每张卡片 ${studentsPerCard} 人</span>
        </div>
    `;
    
    // 显示翻页控件（如果有多张卡片）
    if (images.length > 1) {
        paginationControls.style.display = 'flex';
        populatePageSelector();
        // 更新页面选择器的值为当前页面
        pageSelect.value = currentData.currentPage + 1;
    } else {
        paginationControls.style.display = 'none';
    }
    
    // 显示当前卡片
    displayCurrentCard();
    updatePaginationControls();
    
    // 显示结果区域
    resultSection.style.display = 'block';
    resultSection.scrollIntoView({ behavior: 'smooth' });
}

// 填充页面选择器选项
function populatePageSelector() {
    const currentData = getCurrentData();
    pageSelect.innerHTML = '';
    for (let i = 1; i <= currentData.currentImages.length; i++) {
        const option = document.createElement('option');
        option.value = i;
        option.textContent = i;
        pageSelect.appendChild(option);
    }
}

// 显示当前卡片
function displayCurrentCard() {
    const currentData = getCurrentData();
    if (currentData.currentImages.length === 0) return;
    
    const imagePath = currentData.currentImages[currentData.currentPage];
    // 获取当前页面对应的寝室号
    const dormNumber = currentData.dormInfo && currentData.dormInfo[currentData.currentPage] 
        ? currentData.dormInfo[currentData.currentPage].dorm_number 
        : `group_${currentData.currentPage + 1}`;
    
    currentCard.innerHTML = `
        <div class="card-container">
            <img src="${imagePath}" alt="学生信息卡 ${currentData.currentPage + 1}" class="card-image">
            <div class="card-actions">
                <button onclick="downloadImage('${imagePath}', '${dormNumber}.jpg')" class="download-single-btn">
                    下载当前卡片
                </button>
            </div>
        </div>
    `;
}

// 更新翻页控件状态
function updatePaginationControls() {
    const currentData = getCurrentData();
    if (currentData.currentImages.length <= 1) return;
    
    // 更新按钮状态
    prevBtn.disabled = currentData.currentPage === 0;
    nextBtn.disabled = currentData.currentPage === currentData.currentImages.length - 1;
    
    // 更新页面选择器
    pageSelect.value = currentData.currentPage + 1;
    totalPages.textContent = currentData.currentImages.length;
}

// 下载单个图片
function downloadImage(imagePath, filename) {
    const link = document.createElement('a');
    link.href = imagePath;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// 下载ZIP文件
async function downloadZipFile() {
    const currentData = getCurrentData();
    if (currentData.currentImages.length === 0) {
        showError('没有可下载的图片');
        return;
    }
    
    try {
        // 从第一个图片路径中提取session_id
        const firstImagePath = currentData.currentImages[0];
        console.log('图片路径:', firstImagePath); // 调试信息
        const sessionMatch = firstImagePath.match(/\/static\/output\/([^\/]+)\//); 
        
        if (!sessionMatch) {
            console.error('无法匹配会话ID，图片路径:', firstImagePath);
            showError('无法获取会话ID');
            return;
        }
        
        const sessionId = sessionMatch[1];
        console.log('提取的会话ID:', sessionId); // 调试信息
        
        // 构建ZIP文件名：学院名称-性别-X人寝
        const roomTypeText = getRoomTypeText(currentData.studentsPerCard);
        const collegeName = currentData.collegeName || '未知学院';
        const genderText = currentGender === 'male' ? '男寝' : '女寝';
        const zipFileName = `${collegeName}-${genderText}-${roomTypeText}`;
        
        // 构建下载链接
        const zipUrl = `/download_zip/${sessionId}?filename=${encodeURIComponent(zipFileName)}`;
        
        // 创建下载链接
        const link = document.createElement('a');
        link.href = zipUrl;
        link.download = `${zipFileName}.zip`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        showSuccess('ZIP文件下载已开始');
        
    } catch (error) {
        console.error('下载ZIP文件失败:', error);
        showError('下载ZIP文件失败');
    }
}

// 辅助函数：将数字转换为中文人寝描述
function getRoomTypeText(roomType) {
    const roomTypeMap = {
        2: '二人寝',
        4: '四人寝',
        6: '六人寝',
        8: '八人寝'
    };
    return roomTypeMap[roomType] || `${roomType}人寝`;
}

// 下载所有图片
async function downloadAllImages() {
    const currentData = getCurrentData();
    if (currentData.currentImages.length === 0) return;
    
    for (let i = 0; i < currentData.currentImages.length; i++) {
        const imagePath = currentData.currentImages[i];
        // 获取对应的寝室号
        const dormNumber = currentData.dormInfo && currentData.dormInfo[i] 
            ? currentData.dormInfo[i].dorm_number 
            : `group_${i + 1}`;
        const filename = `${dormNumber}.jpg`;
        
        try {
            const response = await fetch(imagePath);
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            
            const link = document.createElement('a');
            link.href = url;
            link.download = filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            
            window.URL.revokeObjectURL(url);
            
            // 添加延迟避免浏览器阻止多个下载
            if (i < currentData.currentImages.length - 1) {
                await new Promise(resolve => setTimeout(resolve, 500));
            }
        } catch (error) {
            console.error(`下载图片 ${filename} 失败:`, error);
        }
    }
    
    showSuccess(`已下载 ${currentData.currentImages.length} 张信息卡`);
}

// 显示错误信息
function showError(message) {
    // 创建错误提示
    const errorDiv = document.createElement('div');
    errorDiv.className = 'error-message';
    errorDiv.textContent = message;
    
    // 添加到页面顶部
    document.body.insertBefore(errorDiv, document.body.firstChild);
    
    // 3秒后自动移除
    setTimeout(() => {
        if (errorDiv.parentNode) {
            errorDiv.parentNode.removeChild(errorDiv);
        }
    }, 3000);
}

// 显示成功信息
function showSuccess(message) {
    const successDiv = document.createElement('div');
    successDiv.className = 'success-message';
    successDiv.textContent = message;
    
    document.body.insertBefore(successDiv, document.body.firstChild);
    
    setTimeout(() => {
        if (successDiv.parentNode) {
            successDiv.parentNode.removeChild(successDiv);
        }
    }, 3000);
}

// 初始化位置调整功能
// 初始化导航栏
function initializeNavigation() {
    // 寝室类型选择
    const navTabs = document.querySelectorAll('.nav-tab');
    navTabs.forEach(tab => {
        tab.addEventListener('click', function() {
            const roomType = parseInt(this.dataset.roomType);
            switchRoomType(roomType);
        });
    });
    
    // 男女寝室选择
    const genderTabs = document.querySelectorAll('.gender-tab');
    genderTabs.forEach(tab => {
        tab.addEventListener('click', function() {
            const gender = this.dataset.gender;
            switchGender(gender);
        });
    });
}

// 切换性别
function switchGender(gender) {
    currentGender = gender;
    
    // 更新性别选择按钮的激活状态
    const genderTabs = document.querySelectorAll('.gender-tab');
    genderTabs.forEach(tab => {
        tab.classList.remove('active');
        if (tab.dataset.gender === gender) {
            tab.classList.add('active');
        }
    });
    
    // 更新文件上传状态显示
    updateFileUploadStatus();
    
    // 更新结果显示区域
    updateResultsDisplay();
}

// 切换寝室类型
function switchRoomType(roomType) {
    currentRoomType = roomType;
    
    // 更新导航栏激活状态
    const navTabs = document.querySelectorAll('.nav-tab');
    navTabs.forEach(tab => {
        tab.classList.remove('active');
        if (parseInt(tab.dataset.roomType) === roomType) {
            tab.classList.add('active');
        }
    });
    
    // 隐藏所有位置配置区域
    const allConfigs = document.querySelectorAll('.position-config-section');
    allConfigs.forEach(config => {
        config.style.display = 'none';
    });
    
    // 显示当前寝室类型的配置区域
    const currentConfig = document.getElementById(`positionConfigSection${roomType}`);
    if (currentConfig) {
        currentConfig.style.display = 'block';
    }
    
    // 重新初始化当前寝室类型的位置调整
    initializeCurrentRoomTypeAdjustments();
    
    // 更新文件上传状态显示
    updateFileUploadStatus();
    
    // 更新结果显示区域
    updateResultsDisplay();
}

// 初始化当前寝室类型的位置调整
function initializeCurrentRoomTypeAdjustments() {
    // 移除之前的事件监听器
    removeAllEventListeners();
    
    // 为当前寝室类型添加事件监听器
    addCurrentRoomTypeEventListeners();
    
    // 初始化按钮事件
    const applyBtn = document.getElementById(`applyAllBtn_${currentRoomType}`);
    const resetBtn = document.getElementById(`resetAllBtn_${currentRoomType}`);
    
    if (applyBtn) {
        applyBtn.addEventListener('click', applyAllPositionAdjustments);
    }
    
    if (resetBtn) {
        resetBtn.addEventListener('click', resetAllPositionAdjustments);
    }
}

// 移除所有事件监听器
function removeAllEventListeners() {
    // 移除所有寝室类型的输入框事件监听器
    const roomTypes = [2, 3, 4, 5];
    const paramNames = ['PhotoX', 'PhotoY', 'PhotoWidth', 'PhotoHeight', 'TextX', 'TextY', 'FontSize', 'LineSpacing'];
    
    roomTypes.forEach(roomType => {
        const maxPositions = roomType;
        for (let i = 1; i <= maxPositions; i++) {
            paramNames.forEach(paramName => {
                const element = getInputElement(i, paramName, roomType);
                if (element) {
                    // 克隆元素来移除所有事件监听器
                    const newElement = element.cloneNode(true);
                    element.parentNode.replaceChild(newElement, element);
                }
            });
        }
    });
}

// 为当前寝室类型添加事件监听器
function addCurrentRoomTypeEventListeners() {
    for (let i = 1; i <= currentRoomType; i++) {
        const suffix = `_${currentRoomType}`;
        const inputs = [
            document.getElementById(`position${i}PhotoX${suffix}`),
            document.getElementById(`position${i}PhotoY${suffix}`),
            document.getElementById(`position${i}PhotoWidth${suffix}`),
            document.getElementById(`position${i}PhotoHeight${suffix}`),
            document.getElementById(`position${i}TextX${suffix}`),
            document.getElementById(`position${i}TextY${suffix}`),
            document.getElementById(`position${i}FontSize${suffix}`),
            document.getElementById(`position${i}LineSpacing${suffix}`)
        ];
        
        inputs.forEach(input => {
            if (input) {
                input.addEventListener('input', updatePositionAdjustments);
            }
        });
    }
}

function initializePositionAdjustments() {
    // 应用所有调整按钮事件
    if (applyAllBtn) {
        applyAllBtn.addEventListener('click', applyAllPositionAdjustments);
    }
    
    // 重置所有调整按钮事件
    if (resetAllBtn) {
        resetAllBtn.addEventListener('click', resetAllPositionAdjustments);
    }
    
    // 为所有输入框添加实时更新事件
    addInputEventListeners();
}

// 为所有输入框添加实时更新事件
function addInputEventListeners() {
    // 移除旧的事件监听器
    removeAllEventListeners();
    
    // 为当前寝室类型的所有输入框添加事件监听器
    const maxPositions = currentRoomType;
    const paramNames = ['PhotoX', 'PhotoY', 'PhotoWidth', 'PhotoHeight', 'TextX', 'TextY', 'FontSize', 'LineSpacing'];
    
    for (let i = 1; i <= maxPositions; i++) {
        paramNames.forEach(paramName => {
            const element = getInputElement(i, paramName, currentRoomType);
            if (element) {
                element.addEventListener('input', updatePositionAdjustments);
            }
        });
    }
}

// 更新位置调整数据
function updatePositionAdjustments() {
    if (!positionAdjustments[currentRoomType]) {
        return;
    }
    
    for (let i = 1; i <= currentRoomType; i++) {
        const suffix = `_${currentRoomType}`;
        
        const photoX = document.getElementById(`position${i}PhotoX${suffix}`);
        const photoY = document.getElementById(`position${i}PhotoY${suffix}`);
        const photoWidth = document.getElementById(`position${i}PhotoWidth${suffix}`);
        const photoHeight = document.getElementById(`position${i}PhotoHeight${suffix}`);
        const textX = document.getElementById(`position${i}TextX${suffix}`);
        const textY = document.getElementById(`position${i}TextY${suffix}`);
        const fontSize = document.getElementById(`position${i}FontSize${suffix}`);
        const lineSpacing = document.getElementById(`position${i}LineSpacing${suffix}`);
        
        if (photoX && photoY && photoWidth && photoHeight && textX && textY && fontSize && lineSpacing) {
            positionAdjustments[currentRoomType][i] = {
                photo_x: parseInt(photoX.value) || 0,
                photo_y: parseInt(photoY.value) || 0,
                photo_width: parseInt(photoWidth.value) || 0,
                photo_height: parseInt(photoHeight.value) || 0,
                text_x: parseInt(textX.value) || 0,
                text_y: parseInt(textY.value) || 0,
                font_size: parseInt(fontSize.value) || 0,
                line_spacing: parseInt(lineSpacing.value) || 0
            };
        }
    }
}

// 应用所有位置调整
async function applyAllPositionAdjustments() {
    updatePositionAdjustments();
    
    // 如果已经有生成的信息卡，自动重新生成以应用调整
    const currentData = getCurrentData();
    if (currentData.currentImages.length > 0 && currentData.excelFile && currentData.imageFile) {
        showSuccess('正在应用位置调整并重新生成信息卡...');
        await generateCards();
    } else {
        showSuccess('位置调整已保存，生成信息卡时将应用这些设置');
    }
}

// 重置所有位置调整
function resetAllPositionAdjustments() {
    // 重置当前寝室类型的所有输入框
    for (let i = 1; i <= currentRoomType; i++) {
        const suffix = `_${currentRoomType}`;
        const inputs = [
            document.getElementById(`position${i}PhotoX${suffix}`),
            document.getElementById(`position${i}PhotoY${suffix}`),
            document.getElementById(`position${i}PhotoWidth${suffix}`),
            document.getElementById(`position${i}PhotoHeight${suffix}`),
            document.getElementById(`position${i}TextX${suffix}`),
            document.getElementById(`position${i}TextY${suffix}`),
            document.getElementById(`position${i}FontSize${suffix}`),
            document.getElementById(`position${i}LineSpacing${suffix}`)
        ];
        
        inputs.forEach(input => {
            if (input) {
                input.value = '';
            }
        });
    }
    
    // 重置当前寝室类型的调整数据
    if (positionAdjustments[currentRoomType]) {
        for (let i = 1; i <= currentRoomType; i++) {
            positionAdjustments[currentRoomType][i] = {
                photo_x: 0, photo_y: 0, photo_width: 0, photo_height: 0,
                text_x: 0, text_y: 0, font_size: 0, line_spacing: 0
            };
        }
    }
    
    showSuccess('所有位置调整已重置');
}

// 更新文件上传状态显示
// 初始化默认位置参数
function initializeDefaultPositions() {
    // 为所有寝室类型初始化默认位置参数
    [2, 3, 4, 5].forEach(roomType => {
        for (let i = 1; i <= roomType; i++) {
            const suffix = `_${roomType}`;
            
            const photoX = document.getElementById(`position${i}PhotoX${suffix}`);
            const photoY = document.getElementById(`position${i}PhotoY${suffix}`);
            const photoWidth = document.getElementById(`position${i}PhotoWidth${suffix}`);
            const photoHeight = document.getElementById(`position${i}PhotoHeight${suffix}`);
            const textX = document.getElementById(`position${i}TextX${suffix}`);
            const textY = document.getElementById(`position${i}TextY${suffix}`);
            const fontSize = document.getElementById(`position${i}FontSize${suffix}`);
            const lineSpacing = document.getElementById(`position${i}LineSpacing${suffix}`);
            
            // 如果输入框有默认值，则使用默认值更新positionAdjustments
            if (photoX && photoX.value) {
                positionAdjustments[roomType][i].photo_x = parseInt(photoX.value);
            }
            if (photoY && photoY.value) {
                positionAdjustments[roomType][i].photo_y = parseInt(photoY.value);
            }
            if (photoWidth && photoWidth.value) {
                positionAdjustments[roomType][i].photo_width = parseInt(photoWidth.value);
            }
            if (photoHeight && photoHeight.value) {
                positionAdjustments[roomType][i].photo_height = parseInt(photoHeight.value);
            }
            if (textX && textX.value) {
                positionAdjustments[roomType][i].text_x = parseInt(textX.value);
            }
            if (textY && textY.value) {
                positionAdjustments[roomType][i].text_y = parseInt(textY.value);
            }
            if (fontSize && fontSize.value) {
                positionAdjustments[roomType][i].font_size = parseInt(fontSize.value);
            }
            if (lineSpacing && lineSpacing.value) {
                positionAdjustments[roomType][i].line_spacing = parseInt(lineSpacing.value);
            }
        }
    });
}

function updateFileUploadStatus() {
    const currentData = getCurrentData();
    
    // 更新Excel文件状态
    if (currentData.excelFile) {
        excelInfo.innerHTML = `
            <div class="file-info">
                <span class="file-name">${currentData.excelFile.name}</span>
                <span class="file-size">${formatFileSize(currentData.excelFile.size)}</span>
            </div>
        `;
        excelUpload.classList.add('has-file');
    } else {
        excelInfo.innerHTML = '<span class="upload-text">点击或拖拽上传Excel文件</span>';
        excelUpload.classList.remove('has-file');
    }
    
    // 更新图片文件状态
    if (currentData.imageFile) {
        imageInfo.innerHTML = `
            <div class="file-info">
                <span class="file-name">${currentData.imageFile.name}</span>
                <span class="file-size">${formatFileSize(currentData.imageFile.size)}</span>
            </div>
        `;
        imageUpload.classList.add('has-file');
    } else {
        imageInfo.innerHTML = '<span class="upload-text">点击或拖拽上传背景图片</span>';
        imageUpload.classList.remove('has-file');
    }
    
    // 更新生成按钮状态
    checkFilesReady();
}

// 更新结果显示区域
function updateResultsDisplay() {
    const currentData = getCurrentData();
    
    if (currentData.hasResults && currentData.currentImages.length > 0) {
        // 显示当前寝室类型的结果
        displayResults(currentData.currentImages, currentData.totalStudents, currentData.studentsPerCard);
    } else {
        // 隐藏结果区域
        resultSection.style.display = 'none';
    }
}