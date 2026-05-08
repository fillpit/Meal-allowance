/* ==========================================================================
   食堂补贴导入模板自动生成器 - 核心业务逻辑
   特色：账号清洗去重、日期星期精确过滤、大数据量截断预览、原生无依赖Excel编码
   ========================================================================== */

document.addEventListener('DOMContentLoaded', () => {
  
  // ==========================================================================
  // 1. 全局状态变量
  // ==========================================================================
  let parsedAccounts = []; // 清洗去重后的账户列表
  let generatedRecords = []; // 最终生成的笛卡尔积明细数据列表
  let filteredRecords = []; // 经过搜索筛选后的明细数据列表
  let loadedFileName = ''; // 导入的文件名称

  // ==========================================================================
  // 2. DOM 元素获取
  // ==========================================================================
  const bodyEl = document.body;
  const themeToggleBtn = document.getElementById('theme-toggle');
  
  // 卡片1: 人员名单相关
  const tabButtons = document.querySelectorAll('.tab-btn');
  const tabContents = document.querySelectorAll('.tab-content');
  const usersTextarea = document.getElementById('users-input');
  const accountsCountBadge = document.getElementById('accounts-count');
  const dragZone = document.getElementById('drag-zone');
  const fileInput = document.getElementById('file-input');
  const fileInfoBadge = document.getElementById('file-info');
  const clearFileBtn = document.getElementById('clear-file');
  
  // 卡片2: 日期选择相关
  const startDateInput = document.getElementById('start-date');
  const endDateInput = document.getElementById('end-date');
  const weekdaySelector = document.getElementById('weekday-selector');
  const weekdayCheckboxes = document.querySelectorAll('.weekday-checkbox');
  const weekdayItems = document.querySelectorAll('.weekday-item');
  const selectWorkdaysBtn = document.getElementById('select-workdays');
  const selectAllDaysBtn = document.getElementById('select-all-days');
  
  // 卡片3: 补贴金额类型相关
  const subsidyAmountInput = document.getElementById('subsidy-amount');
  const quickAmountChips = document.querySelectorAll('#quick-amounts .chip-btn');
  const typeCards = document.querySelectorAll('.type-card');
  const typeRadios = document.querySelectorAll('.type-radio');
  
  // 卡片4: 预览及数据汇总相关
  const searchInput = document.getElementById('search-input');
  const statUsersVal = document.getElementById('stat-users');
  const statDaysVal = document.getElementById('stat-days');
  const statRecordsVal = document.getElementById('stat-records');
  const statTotalAmountVal = document.getElementById('stat-total-amount');
  
  const tableViewContainer = document.getElementById('table-view-container');
  const emptyStateView = document.getElementById('empty-state-view');
  const previewTableEl = document.getElementById('preview-table-el');
  const previewTbody = document.getElementById('preview-tbody');
  const tableFooterEl = document.getElementById('table-footer-el');
  const rowsInfoText = document.getElementById('rows-info-text');
  const clearAllRecordsBtn = document.getElementById('clear-all-records-btn');
  
  const exportBtn = document.getElementById('export-btn');
  const toastContainer = document.getElementById('toast-container');

  // ==========================================================================
  // 3. 初始化默认值
  // ==========================================================================
  function initDefaults() {
    const today = new Date();
    const todayStr = formatDateToYYYYMMDD(today);
    
    // 默认开始日期为今天
    startDateInput.value = todayStr;
    
    // 默认结束日期为今天往后延 6 天（总计包含今天共 7 天）
    const nextWeek = new Date(today);
    nextWeek.setDate(today.getDate() + 6);
    endDateInput.value = formatDateToYYYYMMDD(nextWeek);
    
    // 初始化日期控件的可选范围（避免手滑选择太久远的时间）
    startDateInput.min = '2020-01-01';
    endDateInput.min = '2020-01-01';
    
    // 默认执行一次重新计算
    triggerRecalculation();
  }

  // ==========================================================================
  // 4. 辅助函数：消息通知 (Toast) 与格式化
  // ==========================================================================
  function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    
    // 动态图标选择
    let iconSvg = '';
    if (type === 'success') {
      iconSvg = `<svg class="toast-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="20 6 9 17 4 12"/></svg>`;
    } else if (type === 'error') {
      iconSvg = `<svg class="toast-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>`;
    } else {
      iconSvg = `<svg class="toast-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>`;
    }
    
    toast.innerHTML = `
      ${iconSvg}
      <span class="toast-message">${message}</span>
      <button class="toast-close">&times;</button>
    `;
    
    toastContainer.appendChild(toast);
    
    // 绑定关闭按钮事件
    toast.querySelector('.toast-close').addEventListener('click', () => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateX(100%)';
      setTimeout(() => toast.remove(), 300);
    });
    
    // 4秒后自动淡出消失
    setTimeout(() => {
      if (toast.parentNode) {
        toast.style.opacity = '0';
        toast.style.transform = 'translateX(100%)';
        setTimeout(() => toast.remove(), 300);
      }
    }, 4000);
  }

  function formatDateToYYYYMMDD(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  function escapeXml(unsafe) {
    return String(unsafe).replace(/[<>&'"]/g, function (c) {
      switch (c) {
        case '<': return '&lt;';
        case '>': return '&gt;';
        case '&': return '&amp;';
        case '\'': return '&apos;';
        case '"': return '&quot;';
        default: return c;
      }
    });
  }

  // ==========================================================================
  // 5. 交互：暗黑/明亮主题切换
  // ==========================================================================
  themeToggleBtn.addEventListener('click', () => {
    if (bodyEl.classList.contains('dark-theme')) {
      bodyEl.classList.remove('dark-theme');
      bodyEl.classList.add('light-theme');
      showToast('已切换至优雅明亮模式 ☀️', 'info');
    } else {
      bodyEl.classList.remove('light-theme');
      bodyEl.classList.add('dark-theme');
      showToast('已切换至极客暗黑模式 🌙', 'info');
    }
  });

  // ==========================================================================
  // 6. 核心逻辑：人员名单清洗解析 (Textarea + Drag/Upload)
  // ==========================================================================
  function parseAndCleanAccounts(rawText) {
    // 按换行、逗号或分号切割
    const items = rawText.split(/[\n,;，；]/);
    const cleaned = [];
    
    items.forEach(item => {
      const trimmed = item.trim();
      // 过滤掉空行、空格，且去除多余字符
      if (trimmed !== '') {
        cleaned.push(trimmed);
      }
    });
    
    // 利用 Set 去除重复人员
    parsedAccounts = [...new Set(cleaned)];
    
    // 更新界面徽章与统计数值
    accountsCountBadge.textContent = `已识别：${parsedAccounts.length} 人`;
    statUsersVal.textContent = parsedAccounts.length;
    
    triggerRecalculation();
  }

  // 监听输入框实时变化
  usersTextarea.addEventListener('input', (e) => {
    parseAndCleanAccounts(e.target.value);
  });

  // Tab 选项卡切换逻辑
  tabButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      tabButtons.forEach(b => b.classList.remove('active'));
      tabContents.forEach(c => c.classList.add('hidden'));
      
      btn.classList.add('active');
      const activeTabId = btn.getAttribute('data-tab');
      document.getElementById(activeTabId).classList.remove('hidden');
    });
  });

  // 处理拖拽上传文件
  dragZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dragZone.classList.add('dragover');
  });

  dragZone.addEventListener('dragleave', () => {
    dragZone.classList.remove('dragover');
  });

  dragZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dragZone.classList.remove('dragover');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      handleFileImport(files[0]);
    }
  });

  dragZone.addEventListener('click', () => {
    fileInput.click();
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      handleFileImport(e.target.files[0]);
    }
  });

  function handleFileImport(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext !== 'txt' && ext !== 'csv') {
      showToast('❌ 格式不支持！请上传 .txt 或 .csv 文本文件', 'error');
      return;
    }
    
    const reader = new FileReader();
    reader.onload = function(e) {
      const text = e.target.result;
      loadedFileName = file.name;
      
      // 同步填充进文本框，方便管理人员看得到和修改
      usersTextarea.value = text;
      
      // 切换显示文件信息徽章
      dragZone.classList.add('hidden');
      fileInfoBadge.classList.remove('hidden');
      fileInfoBadge.querySelector('.file-name').textContent = `📄 ${file.name} (${(file.size/1024).toFixed(1)} KB)`;
      
      // 主动切回粘贴面板，让用户直观地看到导入的内容
      const pasteTabBtn = document.querySelector('[data-tab="paste-tab"]');
      pasteTabBtn.click();
      
      parseAndCleanAccounts(text);
      showToast(`🎉 文件“${file.name}”读取成功，已自动净化！`, 'success');
    };
    reader.readAsText(file);
  }

  // 清除导入的文件
  clearFileBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    fileInput.value = '';
    usersTextarea.value = '';
    loadedFileName = '';
    fileInfoBadge.classList.add('hidden');
    dragZone.classList.remove('hidden');
    parseAndCleanAccounts('');
    showToast('已清除导入的文件与名单。', 'info');
  });

  // ==========================================================================
  // 7. 交互：星期选择器组件
  // ==========================================================================
  
  // 内部 checkbox 状态改变响应（利用 label 默认关联特性，点击任何部分均能精准触发其 change 事件）
  weekdayCheckboxes.forEach(checkbox => {
    checkbox.addEventListener('change', function() {
      const parentLabel = this.closest('.weekday-item');
      if (this.checked) {
        parentLabel.classList.add('checked');
      } else {
        parentLabel.classList.remove('checked');
      }
      triggerRecalculation();
    });
  });

  // 快捷星期动作：工作日
  selectWorkdaysBtn.addEventListener('click', () => {
    weekdayCheckboxes.forEach(box => {
      const val = box.value;
      const isWorkday = (val !== '0' && val !== '6'); // 排除 0(周日) 和 6(周六)
      box.checked = isWorkday;
      
      const label = box.closest('.weekday-item');
      if (isWorkday) {
        label.classList.add('checked');
      } else {
        label.classList.remove('checked');
      }
    });
    showToast('已自动勾选：周一至周五，剔除周末 💼', 'info');
    triggerRecalculation();
  });

  // 快捷星期动作：全选
  selectAllDaysBtn.addEventListener('click', () => {
    weekdayCheckboxes.forEach(box => {
      box.checked = true;
      const label = box.closest('.weekday-item');
      label.classList.add('checked');
    });
    showToast('已全选所有星期 📅', 'info');
    triggerRecalculation();
  });

  // 日期跨度控件变动监听
  startDateInput.addEventListener('change', () => {
    // 联动控制，确保结束日期不早于开始日期
    if (endDateInput.value && endDateInput.value < startDateInput.value) {
      endDateInput.value = startDateInput.value;
    }
    triggerRecalculation();
  });

  endDateInput.addEventListener('change', () => {
    if (startDateInput.value && startDateInput.value > endDateInput.value) {
      startDateInput.value = endDateInput.value;
    }
    triggerRecalculation();
  });

  // ==========================================================================
  // 8. 交互：金额微调、类型选择
  // ==========================================================================
  
  // 绑定快捷金额小泡泡
  quickAmountChips.forEach(chip => {
    chip.addEventListener('click', () => {
      quickAmountChips.forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      
      const val = chip.getAttribute('data-val');
      subsidyAmountInput.value = parseFloat(val).toFixed(2);
      
      triggerRecalculation();
    });
  });

  // 金额数值框手动键入监听
  subsidyAmountInput.addEventListener('input', () => {
    const val = parseFloat(subsidyAmountInput.value);
    
    // 清除快捷泡泡高亮
    quickAmountChips.forEach(c => c.classList.remove('active'));
    
    // 如果碰巧和某个快捷键金额相等，则帮其勾选
    quickAmountChips.forEach(c => {
      if (parseFloat(c.getAttribute('data-val')) === val) {
        c.classList.add('active');
      }
    });
    
    triggerRecalculation();
  });

  // 补贴类型卡片点击交互
  typeCards.forEach(card => {
    card.addEventListener('click', function(e) {
      if (e.target.tagName === 'INPUT') return;
      
      typeCards.forEach(c => c.classList.remove('checked'));
      const radio = this.querySelector('.type-radio');
      radio.checked = true;
      this.classList.add('checked');
      
      triggerRecalculation();
    });
  });

  // 监听单选框的原生变化事件
  typeRadios.forEach(radio => {
    radio.addEventListener('change', function() {
      typeCards.forEach(c => c.classList.remove('checked'));
      if (this.checked) {
        this.closest('.type-card').classList.add('checked');
      }
      triggerRecalculation();
    });
  });

  // ==========================================================================
  // 9. 核心引擎：笛卡尔积数据计算
  // ==========================================================================
  function triggerRecalculation() {
    generatedRecords = [];
    
    const startStr = startDateInput.value;
    const endStr = endDateInput.value;
    const amountVal = parseFloat(subsidyAmountInput.value || 0);
    
    // 获取勾选的星期数组 [0, 1, 2...]
    const activeWeekdays = [];
    weekdayCheckboxes.forEach(box => {
      if (box.checked) {
        activeWeekdays.push(parseInt(box.value));
      }
    });

    // 1. 基本安全拦截
    if (parsedAccounts.length === 0 || !startStr || !endStr || activeWeekdays.length === 0) {
      updatePreviewUI();
      return;
    }

    const startDate = new Date(startStr);
    const endDate = new Date(endStr);
    
    if (startDate > endDate) {
      updatePreviewUI();
      return;
    }

    // 2. 统计计算有效的补贴到账日期
    const validDates = [];
    const loopDate = new Date(startDate);
    
    // 设定安全阀值，最大不允许跨越 366 天，防止卡死
    let daySafetyCount = 0;
    while (loopDate <= endDate && daySafetyCount < 366) {
      const dayOfWeek = loopDate.getDay(); // 0 是周日，1-6 是周一到周六
      if (activeWeekdays.includes(dayOfWeek)) {
        validDates.push(formatDateToYYYYMMDD(loopDate));
      }
      loopDate.setDate(loopDate.getDate() + 1);
      daySafetyCount++;
    }

    // 更新发放天数的汇总数字
    statDaysVal.textContent = validDates.length;

    // 3. 笛卡尔交叉合并生成数据行
    const selectedType = document.querySelector('input[name="subsidy-type"]:checked').value;
    
    parsedAccounts.forEach(account => {
      validDates.forEach(date => {
        generatedRecords.push({
          account: account,
          amount: amountVal,
          type: parseInt(selectedType),
          date: date,
          remarks: '补贴类型：0为累加补贴,1为清零补贴'
        });
      });
    });

    // 默认复制给筛选视图，以进行搜索和呈现
    filteredRecords = [...generatedRecords];
    
    updatePreviewUI();
  }

  // ==========================================================================
  // 10. 渲染：高保真预览表格呈现与局部删除
  // ==========================================================================
  function updatePreviewUI() {
    const totalRecords = filteredRecords.length;
    
    // 更新统计栏
    statRecordsVal.textContent = totalRecords;
    
    // 汇总预计补贴金额总计
    let grandTotal = 0;
    filteredRecords.forEach(r => grandTotal += r.amount);
    statTotalAmountVal.textContent = `￥${grandTotal.toFixed(2)}`;

    // 清空当前的预览表格 DOM
    previewTbody.innerHTML = '';

    // 若无数据，展示空状态
    if (totalRecords === 0) {
      emptyStateView.classList.remove('hidden');
      previewTableEl.classList.add('hidden');
      tableFooterEl.classList.add('hidden');
      exportBtn.disabled = true;
      return;
    }

    emptyStateView.classList.add('hidden');
    previewTableEl.classList.remove('hidden');
    tableFooterEl.classList.remove('hidden');
    exportBtn.disabled = false;

    // 为了防卡顿（食堂名单若有数千人，笛卡尔积轻易上万行，导致DOM卡死）
    // 界面预览仅渲染前 100 行，但后台数组保持完整。
    const maxPreviewLimit = 100;
    const renderLimit = Math.min(totalRecords, maxPreviewLimit);
    
    // 渲染提示文案更新
    if (totalRecords > maxPreviewLimit) {
      rowsInfoText.textContent = `正在预览前 ${maxPreviewLimit} 条明细（全部共 ${totalRecords} 条，点击一键导出下载完整模板）`;
    } else {
      rowsInfoText.textContent = `正在显示全部 ${totalRecords} 条补贴明细`;
    }

    for (let i = 0; i < renderLimit; i++) {
      const rowData = filteredRecords[i];
      const tr = document.createElement('tr');
      
      // 找到该条目在 generatedRecords 数组中的真实全局索引，便于删除
      const realGlobalIndex = generatedRecords.indexOf(rowData);
      
      tr.innerHTML = `
        <td style="font-family: monospace; font-weight: 600;">${escapeXml(rowData.account)}</td>
        <td style="font-weight: 700; color: var(--primary-color);">￥${rowData.amount.toFixed(2)}</td>
        <td>
          <span style="padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; ${
            rowData.type === 0 
              ? 'background: var(--success-bg); color: var(--success-color);' 
              : 'background: var(--danger-bg); color: var(--danger-color);'
          }">
            ${rowData.type === 0 ? '累加 (0)' : '清零 (1)'}
          </span>
        </td>
        <td>${rowData.date}</td>
        <td style="font-size: 11px; color: var(--text-secondary); max-width: 220px; overflow: hidden; text-overflow: ellipsis;" title="${rowData.remarks}">
          ${escapeXml(rowData.remarks)}
        </td>
        <td>
          <button class="action-btn-del" data-index="${realGlobalIndex}" title="在当前预览中移除本条">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <polyline points="3 6 5 6 21 6"/>
              <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>
              <line x1="10" y1="11" x2="10" y2="17"/>
              <line x1="14" y1="11" x2="14" y2="17"/>
            </svg>
          </button>
        </td>
      `;
      
      previewTbody.appendChild(tr);
    }

    // 绑定小垃圾桶点击删除事件
    const delBtns = previewTbody.querySelectorAll('.action-btn-del');
    delBtns.forEach(btn => {
      btn.addEventListener('click', function(e) {
        e.stopPropagation();
        const indexToDel = parseInt(this.getAttribute('data-index'));
        if (indexToDel >= 0 && indexToDel < generatedRecords.length) {
          // 在大数组中移除
          generatedRecords.splice(indexToDel, 1);
          // 在当前筛选数组中也同步移除本对象
          const filteredIndex = filteredRecords.indexOf(generatedRecords[indexToDel]);
          if (filteredIndex > -1) {
            filteredRecords.splice(filteredIndex, 1);
          }
          
          // 执行软计算渲染
          triggerSearchAndFilter();
          showToast('已从当前导出批次中移除了该条记录。', 'info');
        }
      });
    });
  }

  // ==========================================================================
  // 11. 搜索筛选预览明细
  // ==========================================================================
  function triggerSearchAndFilter() {
    const searchVal = searchInput.value.trim().toLowerCase();
    
    if (searchVal === '') {
      filteredRecords = [...generatedRecords];
    } else {
      filteredRecords = generatedRecords.filter(r => {
        return r.account.toLowerCase().includes(searchVal) || r.date.includes(searchVal);
      });
    }
    
    updatePreviewUI();
  }

  // 搜索框键入时，一律执行模糊筛选
  searchInput.addEventListener('input', triggerSearchAndFilter);

  // 清空全部数据操作
  clearAllRecordsBtn.addEventListener('click', () => {
    if (confirm('⚠️ 确定要清空当前生成的所有补贴明细吗？\n（这会同时清除左侧输入的人员名单，请确保有备份）')) {
      usersTextarea.value = '';
      fileInput.value = '';
      loadedFileName = '';
      fileInfoBadge.classList.add('hidden');
      dragZone.classList.remove('hidden');
      
      parsedAccounts = [];
      generatedRecords = [];
      filteredRecords = [];
      
      accountsCountBadge.textContent = '已识别：0 人';
      statUsersVal.textContent = '0';
      statDaysVal.textContent = '0';
      
      updatePreviewUI();
      showToast('已成功清空全部明细与名单。', 'info');
    }
  });

  // ==========================================================================
  // 12. 原生 Excel 导出驱动 (SpreadsheetML)
  // ==========================================================================
  exportBtn.addEventListener('click', () => {
    // 再次双保险校验
    if (generatedRecords.length === 0) {
      showToast('❌ 没有可供导出的明细数据！请检查参数', 'error');
      return;
    }

    try {
      // 禁用按钮并展示加载微动画
      exportBtn.disabled = true;
      exportBtn.innerHTML = `
        <svg class="btn-icon" style="animation: spin 1s infinite linear;" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
          <circle cx="12" cy="12" r="10" stroke-dasharray="32" stroke-dashoffset="16"/>
        </svg>
        正在构建 Excel 数据...
      `;

      // 启动延迟，优化加载视觉，防止直接闪现让管理员以为没反应
      setTimeout(() => {
        const xmlContent = buildSpreadsheetML(generatedRecords);
        
        // 构建 Blob
        const blob = new Blob([xmlContent], { type: 'application/vnd.ms-excel;charset=utf-8;' });
        const downloadUrl = URL.createObjectURL(blob);
        
        // 虚拟点击 A 标签触发浏览器下载
        const downloadLink = document.createElement('a');
        downloadLink.href = downloadUrl;
        
        // 按照规范导出 .xls 兼容格式
        downloadLink.download = '食堂补贴导入模板.xls';
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);
        
        // 释放 Blob URL
        URL.revokeObjectURL(downloadUrl);
        
        // 恢复按钮状态
        exportBtn.disabled = false;
        exportBtn.innerHTML = `
          <svg class="btn-icon animate-bounce" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
            <polyline points="7 10 12 15 17 10"/>
            <line x1="12" y1="15" x2="12" y2="3"/>
          </svg>
          一键生成并导出 Excel 模板
        `;
        
        showToast('🎉 Excel 补贴模板生成并下载成功！请在Excel中打开查看', 'success');
      }, 600);

    } catch (err) {
      console.error(err);
      showToast(`❌ 导出失败: ${err.message}`, 'error');
      exportBtn.disabled = false;
    }
  });

  /* 
   * 构造符合 Excel 2003 XML 规范的 SpreadsheetML 文本
   * 包含：第1行合并标题行、第2行表头底色加粗、精细的实线边框
   */
  function buildSpreadsheetML(records) {
    let xml = `<?xml version="1.0" encoding="utf-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Author>食堂补贴导入模板生成器</Author>
  <LastAuthor>食堂智能管理套件</LastAuthor>
  <Created>${new Date().toISOString()}</Created>
 </DocumentProperties>
 <Styles>
  <!-- 全局默认常规样式 -->
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Center"/>
   <Borders/>
   <Font ss:FontName="宋体" x:CharSet="134" ss:Size="11" ss:Color="#000000"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
  
  <!-- 第1行：标题栏样式：合并、加粗、绿底 -->
  <Style ss:ID="TitleStyle">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="宋体" x:CharSet="134" ss:Size="14" ss:Bold="1" ss:Color="#000000"/>
   <Interior ss:Color="#E2EFDA" ss:Pattern="Solid"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
   </Borders>
  </Style>
  
  <!-- 第2行：表头标题样式：加粗、灰底、黑框 -->
  <Style ss:ID="HeaderStyle">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="宋体" x:CharSet="134" ss:Size="11" ss:Bold="1" ss:Color="#000000"/>
   <Interior ss:Color="#F2F2F2" ss:Pattern="Solid"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
   </Borders>
  </Style>
  
  <!-- 数据行常规文本样式 -->
  <Style ss:ID="DataStyle">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
   </Borders>
  </Style>
  
  <!-- 数据行金额浮点型样式 -->
  <Style ss:ID="DataStyleAmount">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <NumberFormat ss:Format="0.00"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
   </Borders>
  </Style>
 </Styles>
 
 <Worksheet ss:Name="补贴导入">
  <Table ss:ExpandedColumnCount="5" x:FullColumns="1" x:FullRows="1" ss:DefaultRowHeight="18">
   <!-- 定制每一列的宽度 -->
   <Column ss:Width="120"/> <!-- 账户号 -->
   <Column ss:Width="100"/> <!-- 补贴金额 -->
   <Column ss:Width="100"/> <!-- 补贴类型 -->
   <Column ss:Width="120"/> <!-- 到账时间 -->
   <Column ss:Width="260"/> <!-- 备注 -->
   
   <!-- 第1行：标题大单元格合并 -->
   <Row ss:Height="30">
    <Cell ss:MergeAcross="4" ss:StyleID="TitleStyle">
     <Data ss:Type="String">补贴导入模板</Data>
    </Cell>
   </Row>
   
   <!-- 第2行：表头行 -->
   <Row ss:Height="22">
    <Cell ss:StyleID="HeaderStyle"><Data ss:Type="String">账户号</Data></Cell>
    <Cell ss:StyleID="HeaderStyle"><Data ss:Type="String">补贴金额</Data></Cell>
    <Cell ss:StyleID="HeaderStyle"><Data ss:Type="String">补贴类型</Data></Cell>
    <Cell ss:StyleID="HeaderStyle"><Data ss:Type="String">到账时间</Data></Cell>
    <Cell ss:StyleID="HeaderStyle"><Data ss:Type="String">备注</Data></Cell>
   </Row>`;

    // 填充明细行
    records.forEach(r => {
      xml += `
   <Row ss:Height="20">
    <Cell ss:StyleID="DataStyle"><Data ss:Type="String">${escapeXml(r.account)}</Data></Cell>
    <Cell ss:StyleID="DataStyleAmount"><Data ss:Type="Number">${r.amount}</Data></Cell>
    <Cell ss:StyleID="DataStyle"><Data ss:Type="Number">${r.type}</Data></Cell>
    <Cell ss:StyleID="DataStyle"><Data ss:Type="String">${r.date}</Data></Cell>
    <Cell ss:StyleID="DataStyle"><Data ss:Type="String">${escapeXml(r.remarks)}</Data></Cell>
   </Row>`;
    });

    xml += `
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Header x:Margin="0.3"/>
    <Footer x:Margin="0.3"/>
    <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
   </PageSetup>
   <Unsynced/>
   <Selected/>
   <Panes>
    <Pane>
     <Number>3</Number>
     <ActiveRow>1</ActiveRow>
     <ActiveCol>1</ActiveCol>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
</Workbook>`;

    return xml;
  }

  // ==========================================================================
  // 13. 初始化启动
  // ==========================================================================
  initDefaults();
});
