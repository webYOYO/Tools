const tipsArea = document.getElementById('txt');
const numArea = document.getElementById('num');
const progress = document.getElementById('progress');
const progress_box = document.getElementById('progress_box');
const loading = document.getElementById('loader');
const fileInput = document.getElementById('fileInput');
const exportBtn = document.getElementById('exportBtn');

let fileData = "";
let json_arr = [];
let excel_data = [];
let inputFileName = "";
let BMS_NUM = 0;
let current_progress = 0;
// 显示/隐藏加载动画
function loadingAnimation(isShow, progressNum) {
    current_progress = progressNum;
    loading.style.display = isShow ? 'block' : 'none';
    if (progressNum !== null) {
        numArea.textContent = progressNum + '%';
        progress.style.width = progressNum + '%';
    }
}

// 监听文件选择变化
fileInput.addEventListener('change', function (e) {
    // 验证文件类型
    const file = e.target.files[0];
    if (!file) {
        progress_box.style.display = 'none';
        exportBtn.disabled = true;
        return;
    }

    // 检查文件扩展名
    const validExtensions = ['.json'];
    const fileExt = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    inputFileName = file.name.split('.')[0];
    if (validExtensions.includes(fileExt)) {
        // 启用所有需要控制的按钮
        exportBtn.disabled = false;
        progress_box.style.display = 'block';
        // 可选：在控制台显示文件信息
        console.log('已选择文件:', file.name);
        console.log('文件大小:', (file.size / 1024).toFixed(2) + 'KB');
    } else {
        alert('请选择有效的 .json 文件');
        fileInput.value = ''; // 清空选择
    }
    loadingAnimation(0, 0);
    tipsArea.textContent = '';
});

function readFileAsync() {
    return new Promise((resolve, reject) => {
        // 清空显示区域
        tipsArea.textContent = '正在读取文件...';
        // 获取用户选择的文件
        const file = fileInput.files[0];

        // 异常处理
        if (!file) {
            tipsArea.textContent = '⚠️ 请先选择文件';
            return;
        }
        // 禁用按钮
        exportBtn.disabled = true;
        fileInput.disabled = true;
        // 显示加载动画
        loadingAnimation(1, 10);
        // 创建文件阅读器
        const reader = new FileReader();

        reader.onload = (e) => resolve(e.target.result);
        reader.onerror = () => reject('❌ 文件读取失败, 请检查文件内容');
        reader.onabort = () => reject('⚠️ 文件读取被中止');

        // 开始读取文件（以文本格式）
        reader.readAsText(file, 'UTF-8');
    });
}

async function handleFileData() {
    let i = 0;
    BMS_NUM = 0;
    try {
        // 读取文件
        const data = await readFileAsync();
        loadingAnimation(1, 20);
        //分割数据
        json_arr = data.split("<end>");
        // 去除最后一个空字符串
        json_arr.pop();

        if (json_arr.length == 0) {
            throw "文件内容错误，请检查文件内容";
        }
        for (; i < json_arr.length; i++) {
            //每条数据处理平分70%进度条，直到90%
            let progressNum = Math.floor((70 / json_arr.length) * (i + 1)) + 20;
            loadingAnimation(1, progressNum);

            json_arr[i] = JSON.parse(json_arr[i]);
            excel_data[i + 1] = [];
            //处理每条数据
            for (let key in json_arr[i]) {
                //原数据中，每一个温度对应4个电芯温度
                if (key.includes("BMS_T")) {
                    //通过第一组数据确定BMS数量
                    if (i == 0) {
                        BMS_NUM += 4;
                    }
                    for (let temp_idx = 1; temp_idx <= 4; temp_idx++) {
                        excel_data[i + 1].push(json_arr[i][key]);
                    }
                } else {
                    excel_data[i + 1].push(json_arr[i][key]);
                }
            }

        }
        // 下载文件
        downloadFile();
    } catch (error) {
        tipsArea.textContent = `❌ ${current_progress > 20 ? '第' + (i + 1) + '组数据异常' : ''} ${error}`;
        fileInput.disabled = false;
        loadingAnimation(0, null);
        fileInput.value = '';
    }
}

function downloadFile() {
    excel_data[0] = ["时间", "SOC", "Pack电压", "继电器外侧电压", "电流", "绝缘值", "主正接触器状态", "主负接触器状态", "预充接触器状态"];
    //表头增加BMS_NUM个单体电压
    for (let i = 1; i <= BMS_NUM; i++) {
        excel_data[0].push(`Volt${i}`);
    }
    //表头增加BMS_NUM个温度
    for (let i = 1; i <= BMS_NUM; i++) {
        excel_data[0].push(`Temp${i}`);
    }

    loadingAnimation(1, 95);

    // 2. 创建工作簿和工作表
    const wb = XLSX.utils.book_new();  // 新建工作簿
    const ws = XLSX.utils.aoa_to_sheet(excel_data);  // 将数组转换为工作表

    // 3. 设置列宽（可选）
    // ws["!cols"] = [
    //     { wch: 15 },  // 第一列宽度15字符
    //     { wch: 10 },
    //     { wch: 20 },
    //     { wch: 15 }
    // ];

    // 4. 将工作表添加到工作簿
    XLSX.utils.book_append_sheet(wb, ws, "单体电压与温度数据");

    // 5. 导出文件
    XLSX.writeFile(wb, `${inputFileName}.xlsx`, {
        cellDates: true,  // 允许日期格式
        bookType: 'xlsx'  // 指定文件类型
    });
    tipsArea.textContent = '✅处理完成！';
    loadingAnimation(0, 100);

    fileInput.disabled = false;
}

async function exportExcel() {
    // 初始化
    fileData = "";
    json_arr = [];
    excel_data = [];

    handleFileData();
}