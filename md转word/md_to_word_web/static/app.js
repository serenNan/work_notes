// Markdown 转 Word 前端逻辑

class MarkdownConverter {
    constructor() {
        this.currentFile = null;
        this.currentDownloadUrl = null;
        this.initEventListeners();
    }

    initEventListeners() {
        // 获取 DOM 元素
        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');
        const convertBtn = document.getElementById('convertBtn');
        const resetBtn = document.getElementById('resetBtn');
        const downloadBtn = document.getElementById('downloadBtn');
        const newConvertBtn = document.getElementById('newConvertBtn');
        const retryBtn = document.getElementById('retryBtn');
        const clickHere = document.querySelector('.click-here');

        // 拖拽事件
        dropZone.addEventListener('dragover', (e) => this.handleDragOver(e, dropZone));
        dropZone.addEventListener('dragleave', (e) => this.handleDragLeave(e, dropZone));
        dropZone.addEventListener('drop', (e) => this.handleDrop(e, dropZone));

        // 点击选择文件
        clickHere.addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', (e) => this.handleFileSelect(e.target.files));

        // 按钮事件
        convertBtn.addEventListener('click', () => this.startConversion());
        resetBtn.addEventListener('click', () => this.reset());
        downloadBtn.addEventListener('click', () => this.downloadFile());
        newConvertBtn.addEventListener('click', () => this.reset());
        retryBtn.addEventListener('click', () => this.reset());
    }

    handleDragOver(e, dropZone) {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.add('active');
    }

    handleDragLeave(e, dropZone) {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('active');
    }

    handleDrop(e, dropZone) {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('active');

        const files = e.dataTransfer.files;
        this.handleFileSelect(files);
    }

    handleFileSelect(files) {
        if (files.length === 0) return;

        const file = files[0];

        // 验证文件类型
        if (!file.name.endsWith('.md') && !file.name.endsWith('.markdown')) {
            this.showError('只支持 .md 和 .markdown 文件');
            return;
        }

        // 验证文件大小（50MB）
        if (file.size > 50 * 1024 * 1024) {
            this.showError('文件不能超过 50MB');
            return;
        }

        this.currentFile = file;

        // 显示选项面板
        this.showSection('optionsSection');

        // 更新文件信息
        document.querySelector('.drop-text').textContent = `已选择: ${file.name}`;
    }

    async startConversion() {
        if (!this.currentFile) {
            this.showError('请先选择文件');
            return;
        }

        // 显示进度
        this.showSection('progressSection');

        const formData = new FormData();
        formData.append('file', this.currentFile);
        formData.append('generate_toc', document.getElementById('generateToc').checked);
        formData.append('toc_depth', 3);
        formData.append('highlight_style', document.getElementById('highlightStyle').value);

        try {
            const response = await fetch('/api/convert', {
                method: 'POST',
                body: formData
            });

            const data = await response.json();

            if (data.success) {
                this.currentDownloadUrl = data.download_url;
                this.showResult(data.message, data.filename);
            } else {
                this.showError(data.message || '转换失败，请重试');
            }

        } catch (error) {
            console.error('错误:', error);
            this.showError('网络错误，请检查连接');
        }
    }

    downloadFile() {
        if (!this.currentDownloadUrl) {
            this.showError('下载链接丢失，请重新转换');
            return;
        }

        // 创建临时链接下载
        const link = document.createElement('a');
        link.href = this.currentDownloadUrl;
        link.click();

        // 清理临时文件（可选，让服务器自动清理）
        setTimeout(() => {
            const filename = this.currentDownloadUrl.split('/').pop();
            fetch(`/api/cleanup/${filename}`, { method: 'POST' }).catch(() => {});
        }, 1000);
    }

    showResult(message, filename) {
        document.getElementById('resultText').textContent = message;

        // 设置下载按钮
        const downloadBtn = document.getElementById('downloadBtn');
        downloadBtn.href = this.currentDownloadUrl;
        downloadBtn.download = filename;

        this.showSection('resultSection');
    }

    showError(message) {
        document.getElementById('errorText').textContent = message;
        this.showSection('errorSection');
    }

    showSection(sectionId) {
        // 隐藏所有 section
        document.querySelectorAll('[id$="Section"], #optionsSection, #progressSection, #resultSection, #errorSection').forEach(section => {
            section.style.display = 'none';
        });

        // 显示指定 section
        document.getElementById(sectionId).style.display = 'block';
    }

    reset() {
        this.currentFile = null;
        this.currentDownloadUrl = null;

        // 重置表单
        document.getElementById('fileInput').value = '';
        document.getElementById('generateToc').checked = true;
        document.getElementById('highlightStyle').value = 'tango';

        // 重置显示文本
        document.querySelector('.drop-text').textContent = '拖拽 MD 文件到此处';

        // 显示上传区域
        document.getElementById('dropZone').classList.remove('active');
        this.showSection('uploadSection');
    }
}

// 页面加载完成后初始化
document.addEventListener('DOMContentLoaded', () => {
    // 检查后端是否可用
    fetch('/api/health')
        .then(r => r.json())
        .then(data => {
            console.log('后端状态:', data);
        })
        .catch(err => {
            console.warn('无法连接到后端:', err);
            alert('警告：无法连接到后端服务，请确保服务已启动');
        });

    // 初始化应用
    window.app = new MarkdownConverter();
});

// 覆盖默认的显示逻辑，确保初始显示上传区域
function showInitialSection() {
    const uploadSection = document.getElementById('uploadSection');
    const optionsSection = document.getElementById('optionsSection');
    const progressSection = document.getElementById('progressSection');
    const resultSection = document.getElementById('resultSection');
    const errorSection = document.getElementById('errorSection');

    if (uploadSection) uploadSection.style.display = 'block';
    if (optionsSection) optionsSection.style.display = 'none';
    if (progressSection) progressSection.style.display = 'none';
    if (resultSection) resultSection.style.display = 'none';
    if (errorSection) errorSection.style.display = 'none';
}

// 确保初始状态正确
document.addEventListener('DOMContentLoaded', showInitialSection);
