/**
 * Main application controller - handles UI interactions and orchestrates the comparison pipeline.
 */
(function() {
    'use strict';

    // State
    let fileA = null;
    let fileB = null;
    let resultBlob = null;

    // DOM Elements
    const uploadBoxA = document.getElementById('uploadBoxA');
    const uploadBoxB = document.getElementById('uploadBoxB');
    const fileInputA = document.getElementById('fileInputA');
    const fileInputB = document.getElementById('fileInputB');
    const fileInfoA = document.getElementById('fileInfoA');
    const fileInfoB = document.getElementById('fileInfoB');
    const fileNameA = document.getElementById('fileNameA');
    const fileNameB = document.getElementById('fileNameB');
    const fileMetaA = document.getElementById('fileMetaA');
    const fileMetaB = document.getElementById('fileMetaB');
    const removeA = document.getElementById('removeA');
    const removeB = document.getElementById('removeB');
    const authorInput = document.getElementById('authorName');
    const strictModeToggle = document.getElementById('strictMode');
    const btnCompare = document.getElementById('btnCompare');
    const progressSection = document.getElementById('progressSection');
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');
    const previewSection = document.getElementById('previewSection');
    const previewA = document.getElementById('previewA');
    const previewB = document.getElementById('previewB');
    const statsBar = document.getElementById('statsBar');
    const downloadSection = document.getElementById('downloadSection');
    const btnDownload = document.getElementById('btnDownload');
    const downloadHint = document.getElementById('downloadHint');
    const errorSection = document.getElementById('errorSection');
    const errorText = document.getElementById('errorText');

    const SUPPORTED_TYPES = ['.docx', '.pdf', '.xlsx', '.xls'];

    // ==================== File Upload Handling ====================

    function setupUploadBox(box, input, side) {
        box.addEventListener('click', (e) => {
            if (e.target.closest('.btn-remove')) return;
            input.click();
        });

        box.addEventListener('dragover', (e) => {
            e.preventDefault();
            box.classList.add('dragover');
        });

        box.addEventListener('dragleave', () => {
            box.classList.remove('dragover');
        });

        box.addEventListener('drop', (e) => {
            e.preventDefault();
            box.classList.remove('dragover');
            const file = e.dataTransfer.files[0];
            if (file) handleFile(file, side);
        });

        input.addEventListener('change', () => {
            if (input.files[0]) handleFile(input.files[0], side);
        });
    }

    function handleFile(file, side) {
        const ext = '.' + file.name.split('.').pop().toLowerCase();
        if (!SUPPORTED_TYPES.includes(ext)) {
            showError(`不支持的文件格式: ${ext}\n仅支持 .docx / .pdf / .xlsx / .xls 格式`);
            return;
        }

        if (side === 'A') {
            fileA = file;
            showFileInfo('A', file);
        } else {
            fileB = file;
            showFileInfo('B', file);
        }

        updateCompareButton();
        hideError();
    }

    function showFileInfo(side, file) {
        const box = side === 'A' ? uploadBoxA : uploadBoxB;
        const info = side === 'A' ? fileInfoA : fileInfoB;
        const name = side === 'A' ? fileNameA : fileNameB;
        const meta = side === 'A' ? fileMetaA : fileMetaB;
        const content = box.querySelector('.upload-content');

        content.style.display = 'none';
        info.style.display = 'flex';
        box.classList.add('has-file');

        name.textContent = file.name;
        const sizeMB = (file.size / 1024 / 1024).toFixed(2);
        const ext = file.name.split('.').pop().toUpperCase();
        meta.textContent = `${ext} · ${sizeMB} MB`;
    }

    function removeFile(side) {
        const box = side === 'A' ? uploadBoxA : uploadBoxB;
        const info = side === 'A' ? fileInfoA : fileInfoB;
        const content = box.querySelector('.upload-content');
        const input = side === 'A' ? fileInputA : fileInputB;

        if (side === 'A') fileA = null;
        else fileB = null;

        content.style.display = '';
        info.style.display = 'none';
        box.classList.remove('has-file');
        input.value = '';

        updateCompareButton();
    }

    function updateCompareButton() {
        btnCompare.disabled = !(fileA && fileB);
    }

    // ==================== Progress ====================

    function showProgress() {
        progressSection.style.display = '';
        previewSection.style.display = 'none';
        downloadSection.style.display = 'none';
        hideError();
    }

    function updateProgress(percent, text, stepNum) {
        progressFill.style.width = percent + '%';
        progressText.textContent = text;

        for (let i = 1; i <= 5; i++) {
            const step = document.getElementById('step' + i);
            step.classList.remove('active', 'done');
            if (i < stepNum) step.classList.add('done');
            if (i === stepNum) step.classList.add('active');
        }
    }

    function hideProgress() {
        progressSection.style.display = 'none';
    }

    // ==================== Error ====================

    function showError(msg) {
        errorSection.style.display = '';
        errorText.textContent = msg;
    }

    function hideError() {
        errorSection.style.display = 'none';
    }

    // ==================== Preview ====================

    function showPreview(docAContent, docBContent, diffResult) {
        previewSection.style.display = '';

        // Stats bar
        const s = diffResult.stats;
        statsBar.innerHTML = `
            <span class="stat">共 ${s.totalParagraphs} 段</span>
            <span class="stat stat-mod">修改 ${s.modifiedParagraphs} 段</span>
            <span class="stat stat-ins">新增 ${s.insertedParagraphs} 段 (+${s.insertedChars} 字)</span>
            <span class="stat stat-del">删除 ${s.deletedParagraphs} 段 (-${s.deletedChars} 字)</span>
            <span class="stat">未变 ${s.unchangedParagraphs} 段</span>
        `;

        // Preview panels
        previewA.innerHTML = '';
        previewB.innerHTML = '';

        for (const diff of diffResult.diffs) {
            if (diff.type === 'equal') {
                appendPreviewLine(previewA, diff.paraA.text, 'normal');
                appendPreviewLine(previewB, diff.paraB.text, 'normal');
            } else if (diff.type === 'modified') {
                // Show inline diffs
                const lineA = document.createElement('div');
                const lineB = document.createElement('div');
                lineA.style.marginBottom = '8px';
                lineB.style.marginBottom = '8px';

                for (const change of diff.changes) {
                    if (change.removed) {
                        const span = document.createElement('span');
                        span.className = 'diff-del';
                        span.textContent = change.value;
                        lineA.appendChild(span);
                    } else if (change.added) {
                        const span = document.createElement('span');
                        span.className = 'diff-add';
                        span.textContent = change.value;
                        lineB.appendChild(span);
                    } else {
                        const spanA = document.createElement('span');
                        spanA.textContent = change.value;
                        lineA.appendChild(spanA);
                        const spanB = document.createElement('span');
                        spanB.textContent = change.value;
                        lineB.appendChild(spanB);
                    }
                }

                previewA.appendChild(lineA);
                previewB.appendChild(lineB);
            } else if (diff.type === 'deleted') {
                appendPreviewLine(previewA, diff.paraA.text, 'deleted');
                appendPreviewLine(previewB, '', 'placeholder');
            } else if (diff.type === 'inserted') {
                appendPreviewLine(previewA, '', 'placeholder');
                appendPreviewLine(previewB, diff.paraB.text, 'inserted');
            }
        }

        // Warning
        if (diffResult.warning) {
            const warn = document.createElement('div');
            warn.style.cssText = 'padding:10px;background:#fdf6ec;border:1px solid #e6a23c;border-radius:4px;color:#e6a23c;margin-top:12px;font-size:13px;';
            warn.textContent = '⚠ ' + diffResult.warning;
            previewSection.appendChild(warn);
        }
    }

    function appendPreviewLine(container, text, type) {
        const div = document.createElement('div');
        div.style.marginBottom = '8px';
        div.style.minHeight = '20px';

        if (type === 'deleted') {
            div.style.background = '#ffebe9';
            div.style.padding = '2px 4px';
            div.style.borderRadius = '3px';
            div.style.textDecoration = 'line-through';
            div.style.color = '#cf222e';
        } else if (type === 'inserted') {
            div.style.background = '#e6ffec';
            div.style.padding = '2px 4px';
            div.style.borderRadius = '3px';
            div.style.color = '#1a7f37';
        } else if (type === 'placeholder') {
            div.style.background = '#f5f5f5';
            div.style.padding = '2px 4px';
            div.style.borderRadius = '3px';
        }

        div.textContent = text;
        container.appendChild(div);
    }

    // ==================== Main Comparison Pipeline ====================

    async function startComparison() {
        showProgress();
        btnCompare.disabled = true;
        resultBlob = null;

        try {
            // Step 1: Parse Document A
            updateProgress(10, '正在解析原文档...', 1);
            const docAContent = await Parsers.parse(fileA, (msg) => {
                updateProgress(15, msg, 1);
            });
            updateProgress(30, '原文档解析完成', 1);

            // Step 2: Parse Document B
            updateProgress(35, '正在解析修改后文档...', 2);
            const docBContent = await Parsers.parse(fileB, (msg) => {
                updateProgress(40, msg, 2);
            });
            updateProgress(55, '修改后文档解析完成', 2);

            // Step 3: OCR info
            if (docAContent.metadata.ocrUsed || docBContent.metadata.ocrUsed) {
                updateProgress(60, 'OCR处理已完成', 3);
            } else {
                updateProgress(60, '无需OCR处理', 3);
            }

            // Step 4: Diff
            updateProgress(65, '正在对比文档内容...', 4);
            const strictMode = strictModeToggle.checked;
            const diffResult = Differ.compare(docAContent, docBContent, strictMode);
            updateProgress(80, `对比完成，发现 ${diffResult.stats.modifiedParagraphs + diffResult.stats.insertedParagraphs + diffResult.stats.deletedParagraphs} 处差异`, 4);

            // Show preview
            showPreview(docAContent, docBContent, diffResult);

            // Step 5: Generate DOCX
            updateProgress(85, '正在生成修订文档...', 5);
            const author = authorInput.value.trim() || '作者';
            resultBlob = await Generator.generate(docAContent, docBContent, diffResult, author);

            // Validate
            const isValid = await Generator.validate(resultBlob);
            if (!isValid) {
                showError('生成的文档验证失败，文件可能无法正常打开。请尝试重新对比。');
            }

            updateProgress(100, '修订文档生成完成！', 5);

            // Show download
            downloadSection.style.display = '';
            const totalChanges = diffResult.stats.modifiedParagraphs +
                                diffResult.stats.insertedParagraphs +
                                diffResult.stats.deletedParagraphs;
            downloadHint.textContent = `共 ${totalChanges} 处修订 · 修订作者: ${author}`;

            // Auto-hide progress after a moment
            setTimeout(() => {
                hideProgress();
            }, 1500);

        } catch (error) {
            hideProgress();
            showError(error.message || '处理过程中发生未知错误');
            console.error('Comparison error:', error);
        } finally {
            btnCompare.disabled = !(fileA && fileB);
        }
    }

    // ==================== Download ====================

    function downloadResult() {
        if (!resultBlob) return;

        const nameA = fileA.name.replace(/\.[^.]+$/, '');
        const nameB = fileB.name.replace(/\.[^.]+$/, '');
        const filename = `对比结果_${nameA}_vs_${nameB}.docx`;

        const url = URL.createObjectURL(resultBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    // ==================== Init ====================

    function init() {
        setupUploadBox(uploadBoxA, fileInputA, 'A');
        setupUploadBox(uploadBoxB, fileInputB, 'B');

        removeA.addEventListener('click', (e) => {
            e.stopPropagation();
            removeFile('A');
        });
        removeB.addEventListener('click', (e) => {
            e.stopPropagation();
            removeFile('B');
        });

        btnCompare.addEventListener('click', startComparison);
        btnDownload.addEventListener('click', downloadResult);
    }

    // Wait for DOM
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
