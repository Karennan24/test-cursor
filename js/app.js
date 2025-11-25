(function () {
	"use strict";

	// DOM元素（延迟获取，确保DOM已加载）
	let revenueInput, refundInput, revenuePreview, refundPreview;
	let revenueUploadStatus, refundUploadStatus, revenueValidBadge, refundValidBadge;
	let onlyIssues, fieldCheckboxes, selectAllFields, deselectAllFields;
	let exportProblemsBtn, exportIssuesBtn, goAnalysisBtn;
	let issueSummary, revenueIssueSummary, refundIssueSummary;
	
	// 获取DOM元素的函数
	function getDOMElements() {
		revenueInput = document.getElementById("revenueFile");
		refundInput = document.getElementById("refundFile");
		revenuePreview = document.getElementById("revenuePreview");
		refundPreview = document.getElementById("refundPreview");
		revenueUploadStatus = document.getElementById("revenueUploadStatus");
		refundUploadStatus = document.getElementById("refundUploadStatus");
		revenueValidBadge = document.getElementById("revenueValidBadge");
		refundValidBadge = document.getElementById("refundValidBadge");
		onlyIssues = document.getElementById("onlyIssues");
		fieldCheckboxes = document.getElementById("fieldCheckboxes");
		selectAllFields = document.getElementById("selectAllFields");
		deselectAllFields = document.getElementById("deselectAllFields");
		exportProblemsBtn = document.getElementById("exportProblems");
		exportIssuesBtn = document.getElementById("exportIssues");
		goAnalysisBtn = document.getElementById("goAnalysis");
		issueSummary = document.getElementById("issueSummary");
		revenueIssueSummary = document.getElementById("revenueIssueSummary");
		refundIssueSummary = document.getElementById("refundIssueSummary");
	}

	let revenueState = { headers: [], checked: [], issues: [], raw: [] };
	let refundState = { headers: [], checked: [], issues: [], raw: [] };
	let selectedFields = new Set();
	let singleFieldFilter = null; // 单字段筛选（从问题说明点击）
	
	// 编辑状态管理
	let editedData = {
		revenue: new Map(), // key: rowIndex, value: { field: newValue }
		refund: new Map()
	};
	let isEditMode = false; // 是否处于编辑模式

	function getExt(name) {
		const m = String(name || "").toLowerCase().match(/\.([a-z0-9]+)$/);
		return m ? m[1] : "";
	}

	function loadScriptOnce(src) {
		return new Promise((resolve, reject) => {
			// 已存在同src脚本则直接resolve
			const exists = Array.from(document.scripts).some(s => {
				if (s.src) {
					// 检查完整URL或相对路径
					return s.src.includes(src) || s.src.endsWith(src.split('/').pop());
				}
				return false;
			});
			if (exists) { 
				console.log("脚本已存在，跳过加载:", src);
				resolve(); 
				return; 
			}
			const s = document.createElement("script");
			s.src = src;
			s.async = true;
			s.crossOrigin = "anonymous"; // 添加跨域属性，可能有助于某些CDN
			s.onload = () => {
				console.log("脚本加载成功:", src);
				resolve();
			};
			s.onerror = (e) => {
				console.warn("脚本加载失败:", src, e);
				reject(e);
			};
			document.head.appendChild(s);
		});
	}

	// 确保第三方库可用（CDN异常时尝试备用源）
	function ensureLibs() {
		const statusEl = document.getElementById("libStatus");
		function showWarn(msg) {
			if (!statusEl) return;
			statusEl.style.display = "block";
			statusEl.innerHTML = `<div style="color:#f59e0b;">${msg}</div>`;
		}
		const tasks = [];
		// 优先尝试本地 vendor，再回退多个 CDN（国内外）
		if (!(window && window.XLSX)) {
			tasks.push(
				loadScriptOnce("vendor/xlsx.full.min.js")
					.catch(() => loadScriptOnce("https://cdn.bootcdn.net/ajax/libs/xlsx/0.19.3/xlsx.full.min.js"))
					.catch(() => loadScriptOnce("https://cdn.staticfile.org/xlsx/0.19.3/xlsx.full.min.js"))
					.catch(() => loadScriptOnce("https://lf3-cdn-tos.bytecdntp.com/cdn/expire-1-M/xlsx/0.19.3/xlsx.full.min.js"))
					.catch(() => loadScriptOnce("https://cdn.jsdelivr.net/npm/xlsx@0.19.3/dist/xlsx.full.min.js"))
					.catch(() => loadScriptOnce("https://unpkg.com/xlsx@0.19.3/dist/xlsx.full.min.js"))
			);
		}
		if (!(window && window.saveAs)) {
			tasks.push(
				loadScriptOnce("vendor/FileSaver.min.js")
					.catch(() => loadScriptOnce("https://cdn.bootcdn.net/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"))
					.catch(() => loadScriptOnce("https://cdn.staticfile.org/FileSaver.js/2.0.5/FileSaver.min.js"))
					.catch(() => loadScriptOnce("https://lf3-cdn-tos.bytecdntp.com/cdn/expire-1-M/FileSaver.js/2.0.5/FileSaver.min.js"))
					.catch(() => loadScriptOnce("https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js"))
					.catch(() => loadScriptOnce("https://unpkg.com/file-saver@2.0.5/dist/FileSaver.min.js"))
			);
		}
		return Promise.all(tasks).then(() => {
			if (!(window && window.XLSX)) {
				// 仅警告，不抛错；CSV 将走内置解析
				showWarn("未能加载 Excel 解析库（XLSX）。xlsx/xls 无法解析，但可上传 CSV 使用全部功能。");
				console.warn("XLSX missing: 将使用 CSV 解析回退（若上传CSV）");
			}
			if (!(window && window.saveAs)) {
				showWarn("未能加载文件导出库（FileSaver）。导出为Excel不可用，但不影响校验与分析。");
				console.warn("FileSaver missing: 导出受限");
			}
		}).catch((e) => {
			console.error("依赖库加载失败：", e);
			// 不再抛错，允许CSV继续运行
		});
	}

	// 统一使用 ArrayBuffer 读取，交由 SheetJS 自动识别（xls/xlsx/csv）
	function readFileSmart(file) {
		return new Promise((resolve, reject) => {
			const reader = new FileReader();
			reader.onerror = (e) => reject(e);
			reader.onload = (e) => resolve(new Uint8Array(e.target.result));
			reader.readAsArrayBuffer(file);
		});
	}

	function decodeTextSmart(u8) {
		// BOM 检测
		if (u8.length >= 3 && u8[0] === 0xef && u8[1] === 0xbb && u8[2] === 0xbf) {
			return new TextDecoder("utf-8").decode(u8.subarray(3));
		}
		// 先用 UTF-8
		let utf8Text = "";
		try { utf8Text = new TextDecoder("utf-8", { fatal: false }).decode(u8); } catch {}
		// 统计 � 字符比例判断是否乱码
		const bad = (utf8Text.match(/\uFFFD/g) || []).length;
		const ratio = bad / Math.max(utf8Text.length, 1);
		if (ratio < 0.02 && utf8Text.length) return utf8Text;
		// 尝试 GBK
		try {
			// 大多数现代浏览器支持 gbk
			const gbk = new TextDecoder("gbk");
			return gbk.decode(u8);
		} catch {
			return utf8Text; // 回退
		}
	}

	function readFileTextSmart(file) {
		return readFileSmart(file).then(decodeTextSmart);
	}

	function looksLikeXlsx(u8) {
		// ZIP 文件头 "PK" -> xlsx 是zip容器
		return u8 && u8.length >= 2 && u8[0] === 0x50 && u8[1] === 0x4b;
	}

	// 简单CSV解析（自动识别分隔符：逗号/分号/Tab；支持引号）
	function parseCsv(text) {
		// 选择分隔符
		const firstLine = (text.split(/\r?\n/).find(l => l.trim().length > 0) || "");
		const cand = [",", ";", "\t"];
		let delim = ",";
		let maxCnt = -1;
		cand.forEach(d => {
			const c = (firstLine.match(new RegExp("\\" + d, "g")) || []).length;
			if (c > maxCnt) { maxCnt = c; delim = d; }
		});
		const rows = [];
		let i = 0, cur = "", inQuotes = false, row = [];
		function pushCell() { row.push(cur.trim()); cur = ""; }
		function pushRow() {
			// 过滤完全空的行（所有单元格都是空或空白）
			const hasContent = row.some(cell => cell.length > 0);
			if (hasContent) {
				rows.push(row);
			}
			row = [];
		}
		while (i < text.length) {
			const ch = text[i];
			if (inQuotes) {
				if (ch === '"') {
					if (text[i + 1] === '"') { cur += '"'; i += 2; continue; }
					inQuotes = false; i++; continue;
				}
				cur += ch; i++; continue;
			}
			if (ch === '"') { inQuotes = true; i++; continue; }
			if (ch === delim) { pushCell(); i++; continue; }
			if (ch === "\r") { i++; continue; }
			if (ch === "\n") { pushCell(); pushRow(); i++; continue; }
			cur += ch; i++;
		}
		pushCell(); 
		// 最后一行也要检查是否为空
		if (row.some(cell => cell.length > 0)) {
			rows.push(row);
		}
		if (!rows.length) return [];
		const headers = rows[0];
		const data = rows.slice(1)
			.map(r => {
				const o = {};
				headers.forEach((h, idx) => { o[h] = (r[idx] ?? "").trim(); });
				return o;
			})
			// 再次过滤：如果某行的所有字段都为空，则排除
			.filter(r => Object.values(r).some(v => String(v).trim().length > 0));
		return data.slice(0, 1000);
	}

	function parseWorkbook(u8) {
		let wb = null;
		// 优先 array
		try {
			wb = XLSX.read(u8, { type: "array" });
		} catch (e1) {
			// 回退：转为 binary string 解析
			try {
				let binary = "";
				const chunkSize = 0x8000;
				for (let i = 0; i < u8.length; i += chunkSize) {
					const chunk = u8.subarray(i, i + chunkSize);
					binary += String.fromCharCode.apply(null, chunk);
				}
				wb = XLSX.read(binary, { type: "binary" });
			} catch (e2) {
				// 最后回退：按字符串
				try {
					const decoder = new TextDecoder("utf-8", { fatal: false });
					const str = decoder.decode(u8);
					wb = XLSX.read(str, { type: "string" });
				} catch (e3) {
					console.error("Workbook 解析失败:", e1, e2, e3);
					throw e3 || e2 || e1;
				}
			}
		}
		const sheetName = wb.SheetNames[0];
		const sheet = wb.Sheets[sheetName];
		// defval 确保空单元格返回空串，避免undefined
		const json = XLSX.utils.sheet_to_json(sheet, { raw: true, defval: "" });
		return json.slice(0, 1000);
	}

	function renderTable(container, state, selectedFields) {
		container.innerHTML = "";
		if (!state.checked.length) return;
		let headers = state.checked[0].headers;
		
		// 如果设置了单字段筛选，只显示该字段
		if (singleFieldFilter) {
			headers = headers.filter(h => h === singleFieldFilter);
			if (headers.length === 0) {
				container.innerHTML = `<div style="padding:20px;text-align:center;color:var(--muted);">字段"${singleFieldFilter}"不存在</div>`;
				return;
			}
		}
		// 如果选择了字段，只显示选中的字段（且selectedFields不为空）
		else if (selectedFields && selectedFields.size > 0) {
			headers = headers.filter(h => selectedFields.has(h));
			if (headers.length === 0) {
				container.innerHTML = `<div style="padding:20px;text-align:center;color:var(--muted);">请至少选择一个字段</div>`;
				return;
			}
		}
		
		// 筛选
		const rowsToRender = onlyIssues.checked
			? state.checked.filter((r) => r.hasIssue)
			: state.checked;

		const table = document.createElement("table");

		const thead = document.createElement("thead");
		const trh = document.createElement("tr");
		headers.forEach((h) => {
			const th = document.createElement("th");
			th.textContent = h;
			// 如果字段被选中，高亮显示
			if (selectedFields && selectedFields.has(h)) th.classList.add("hl-col");
			trh.appendChild(th);
		});
		thead.appendChild(trh);
		table.appendChild(thead);

		const tbody = document.createElement("tbody");
		const isRevenue = container === revenuePreview;
		const currentState = isRevenue ? revenueState : refundState;
		const editKey = isRevenue ? 'revenue' : 'refund';
		
		rowsToRender.forEach((item, renderIndex) => {
			const tr = document.createElement("tr");
			const originalIndex = currentState.checked.indexOf(item);
			
			if (item.hasIssue) {
				tr.classList.add("issue");
				// 添加问题提示
				tr.title = item.reasons.join("；");
			}
			
			headers.forEach((h) => {
				const td = document.createElement("td");
				const editKeyFull = `${originalIndex}_${h}`;
				const editedValue = editedData[editKey].get(editKeyFull);
				const v = editedValue !== undefined ? editedValue : (item.row[h] == null ? "" : String(item.row[h]));
				
				if (isEditMode) {
					// 编辑模式：使用可编辑的input
					const input = document.createElement("input");
					input.type = "text";
					input.value = v;
					input.className = "table-edit-input";
					input.dataset.rowIndex = originalIndex;
					input.dataset.field = h;
					input.dataset.tableType = editKey;
					
					// 如果该值被编辑过，添加编辑标记
					if (editedValue !== undefined) {
						input.classList.add("edited");
					}
					
					// 监听输入变化
					input.addEventListener("change", (e) => {
						const rowIdx = parseInt(e.target.dataset.rowIndex);
						const field = e.target.dataset.field;
						const tableType = e.target.dataset.tableType;
						const newValue = e.target.value;
						const editKeyFull = `${rowIdx}_${field}`;
						
						// 保存编辑值
						if (newValue !== state.checked[rowIdx].row[field]) {
							editedData[tableType].set(editKeyFull, newValue);
							e.target.classList.add("edited");
						} else {
							editedData[tableType].delete(editKeyFull);
							e.target.classList.remove("edited");
						}
						
						// 更新确认按钮状态
						updateConfirmButtonState();
					});
					
					td.appendChild(input);
				} else {
					// 只读模式：显示文本
					td.textContent = v;
					if (editedValue !== undefined) {
						td.classList.add("edited-cell");
						td.title = "已编辑";
					}
				}
				
				// 如果字段被选中，高亮显示
				if (selectedFields && selectedFields.has(h)) td.classList.add("hl-col");
				tr.appendChild(td);
			});
			tbody.appendChild(tr);
		});
		table.appendChild(tbody);
		container.appendChild(table);
	}

	function updateFieldCheckboxes() {
		const headerSet = new Set([
			...(revenueState.headers || []),
			...(refundState.headers || []),
		]);
		const headers = Array.from(headerSet).sort();
		fieldCheckboxes.innerHTML = headers.map(h => 
			`<label><input type="checkbox" value="${h}" ${selectedFields.has(h) ? 'checked' : ''}> <span>${h}</span></label>`
		).join("");
		
		// 绑定change事件
		fieldCheckboxes.querySelectorAll('input[type="checkbox"]').forEach(cb => {
			cb.addEventListener('change', () => {
				if (cb.checked) {
					selectedFields.add(cb.value);
				} else {
					selectedFields.delete(cb.value);
				}
				renderTable(revenuePreview, revenueState, selectedFields);
				renderTable(refundPreview, refundState, selectedFields);
			});
		});
	}
	
	// 全局函数：按字段筛选预览
	window.filterByField = function(fieldName, tableType) {
		singleFieldFilter = fieldName;
		if (tableType === 'revenue') {
			renderTable(revenuePreview, revenueState, selectedFields);
		} else if (tableType === 'refund') {
			renderTable(refundPreview, refundState, selectedFields);
		}
		// 显示清除筛选按钮
		showClearFieldFilter();
	};
	
	function showClearFieldFilter() {
		// 检查是否已存在清除按钮
		let clearBtn = document.getElementById('clearFieldFilter');
		if (!clearBtn) {
			clearBtn = document.createElement('button');
			clearBtn.id = 'clearFieldFilter';
			clearBtn.className = 'btn-small';
			clearBtn.textContent = '清除字段筛选';
			clearBtn.style.marginLeft = '8px';
			clearBtn.onclick = () => {
				singleFieldFilter = null;
				renderTable(revenuePreview, revenueState, selectedFields);
				renderTable(refundPreview, refundState, selectedFields);
				clearBtn.remove();
			};
			// 插入到工具栏
			const toolbar = document.querySelector('.toolbar');
			if (toolbar) {
				toolbar.appendChild(clearBtn);
			}
		}
	}

	// 字段选择器折叠/展开功能
	const fieldSelectorToggle = document.getElementById("fieldSelectorToggle");
	const fieldCheckboxesContainer = document.getElementById("fieldCheckboxesContainer");
	const fieldSelectorWrapper = document.querySelector(".field-selector-wrapper");
	
	if (fieldSelectorToggle && fieldCheckboxesContainer && fieldSelectorWrapper) {
		fieldSelectorToggle.addEventListener("click", () => {
			const isExpanded = fieldCheckboxesContainer.style.display !== "none";
			if (isExpanded) {
				fieldCheckboxesContainer.style.display = "none";
				fieldSelectorWrapper.classList.remove("expanded");
			} else {
				fieldCheckboxesContainer.style.display = "block";
				fieldSelectorWrapper.classList.add("expanded");
			}
		});
	}

	function updateBadges() {
		// 营收
		if (!revenueState.checked.length) {
			revenueValidBadge.textContent = "待校验";
			revenueValidBadge.className = "badge";
		} else if (revenueState.issues.length === 0) {
			revenueValidBadge.textContent = "数据可用";
			revenueValidBadge.className = "badge ok";
		} else {
			revenueValidBadge.textContent = `发现问题 ${revenueState.issues.length}`;
			revenueValidBadge.className = "badge warn";
		}
		// 退费
		if (!refundState.checked.length) {
			refundValidBadge.textContent = "待校验";
			refundValidBadge.className = "badge";
		} else if (refundState.issues.length === 0) {
			refundValidBadge.textContent = "数据可用";
			refundValidBadge.className = "badge ok";
		} else {
			refundValidBadge.textContent = `发现问题 ${refundState.issues.length}`;
			refundValidBadge.className = "badge warn";
		}
		
		// 更新问题说明
		updateIssueSummary();
	}
	
	function updateIssueSummary() {
		const hasIssues = revenueState.issues.length > 0 || refundState.issues.length > 0;
		if (!hasIssues) {
			issueSummary.style.display = "none";
			return;
		}
		
		issueSummary.style.display = "block";
		
		// 营收表问题说明
		if (revenueState.issues.length > 0) {
			const issueCounts = {};
			const fieldMap = {}; // 字段名映射
			revenueState.issues.forEach(issue => {
				issue.reasons.forEach(reason => {
					// 提取字段名（从【字段名】中提取）
					const match = reason.match(/【([^】]+)】/);
					const fieldName = match ? match[1] : null;
					const key = reason.split('】')[0] || reason;
					issueCounts[key] = (issueCounts[key] || 0) + 1;
					if (fieldName && !fieldMap[key]) {
						fieldMap[key] = fieldName;
					}
				});
			});
			
			let html = '<h4>营收表问题汇总：</h4>';
			Object.entries(issueCounts).forEach(([type, count]) => {
				const fieldName = fieldMap[type];
				const clickable = fieldName ? `style="cursor:pointer;color:var(--brand);text-decoration:underline;" onclick="filterByField('${fieldName}', 'revenue')"` : '';
				html += `<div class="issue-item"><span class="issue-row" ${clickable}>${type}</span>：共 ${count} 条问题</div>`;
			});
			html += '<details style="margin-top:8px;"><summary style="cursor:pointer;color:var(--muted);font-size:12px;">查看详细问题</summary><div style="margin-top:8px;">';
			revenueState.issues.slice(0, 20).forEach(issue => {
				html += `<div class="issue-item"><span class="issue-row">第${issue.index + 1}行</span><div class="issue-reasons">${issue.reasons.join('；')}</div></div>`;
			});
			if (revenueState.issues.length > 20) {
				html += `<div style="color:var(--muted);font-size:12px;margin-top:8px;">...还有 ${revenueState.issues.length - 20} 条问题</div>`;
			}
			html += '</div></details>';
			revenueIssueSummary.innerHTML = html;
		} else {
			revenueIssueSummary.innerHTML = '';
		}
		
		// 退费表问题说明
		if (refundState.issues.length > 0) {
			const issueCounts = {};
			const fieldMap = {}; // 字段名映射
			refundState.issues.forEach(issue => {
				issue.reasons.forEach(reason => {
					// 提取字段名（从【字段名】中提取）
					const match = reason.match(/【([^】]+)】/);
					const fieldName = match ? match[1] : null;
					const key = reason.split('】')[0] || reason;
					issueCounts[key] = (issueCounts[key] || 0) + 1;
					if (fieldName && !fieldMap[key]) {
						fieldMap[key] = fieldName;
					}
				});
			});
			
			let html = '<h4>退费表问题汇总：</h4>';
			Object.entries(issueCounts).forEach(([type, count]) => {
				const fieldName = fieldMap[type];
				const clickable = fieldName ? `style="cursor:pointer;color:var(--brand);text-decoration:underline;" onclick="filterByField('${fieldName}', 'refund')"` : '';
				html += `<div class="issue-item"><span class="issue-row" ${clickable}>${type}</span>：共 ${count} 条问题</div>`;
			});
			html += '<details style="margin-top:8px;"><summary style="cursor:pointer;color:var(--muted);font-size:12px;">查看详细问题</summary><div style="margin-top:8px;">';
			refundState.issues.slice(0, 20).forEach(issue => {
				html += `<div class="issue-item"><span class="issue-row">第${issue.index + 1}行</span><div class="issue-reasons">${issue.reasons.join('；')}</div></div>`;
			});
			if (refundState.issues.length > 20) {
				html += `<div style="color:var(--muted);font-size:12px;margin-top:8px;">...还有 ${refundState.issues.length - 20} 条问题</div>`;
			}
			html += '</div></details>';
			refundIssueSummary.innerHTML = html;
		} else {
			refundIssueSummary.innerHTML = '';
		}
	}

	// 显示进度
	function showProgress(type, progress) {
		const progressBar = document.getElementById(`${type}Progress`);
		const progressFill = document.getElementById(`${type}ProgressFill`);
		if (progressBar && progressFill) {
			progressBar.style.display = "block";
			progressFill.style.width = progress + "%";
		}
	}

	// 隐藏进度
	function hideProgress(type) {
		const progressBar = document.getElementById(`${type}Progress`);
		if (progressBar) {
			progressBar.style.display = "none";
		}
	}

	// 更新数据统计
	function updateDataStats(type, totalRows, issueCount) {
		const statsDiv = document.getElementById(`${type}Stats`);
		const totalRowsEl = document.getElementById(`${type}TotalRows`);
		const issueCountEl = document.getElementById(`${type}IssueCount`);
		
		if (statsDiv && totalRows > 0) {
			statsDiv.style.display = "flex";
			if (totalRowsEl) totalRowsEl.textContent = totalRows;
			if (issueCountEl) issueCountEl.textContent = issueCount;
		}
	}

	function onRevenueFile(file) {
		if (!file) return;
		revenueUploadStatus.textContent = `已选择：${file.name}`;
		showProgress("revenue", 20);
		const ext = getExt(file.name);
		ensureLibs().then(() => {
			showProgress("revenue", 50);
			// 若XLSX仍不可用，且为CSV，则走内置CSV解析快速通道
			if (!(window && window.XLSX)) {
				if (ext === "csv") {
					// 先检查文件头是否为ZIP（实际是xlsx被误命名为csv）
					return readFileSmart(file).then((u8) => {
						if (looksLikeXlsx(u8)) {
							alert("检测到该CSV文件实际是Excel工作簿（xlsx）。当前未加载Excel解析库，无法直接解析，请将文件正确保存为CSV，或把 xlsx.full.min.js 放入 vendor/ 目录后重试。");
							throw new Error("CSV masqueraded XLSX without library");
						}
						return decodeTextSmart(u8);
					}).then(parseCsv);
				}
				alert("当前环境未加载 Excel 解析库，无法直接解析 xlsx/xls。请将该文件另存为 CSV 后上传，或将 xlsx.full.min.js 放入 vendor/ 目录。");
				throw new Error("缺少XLSX库且文件并非CSV，无法解析。");
			}
			return readFileSmart(file).then(parseWorkbook);
		}).then((rows) => {
			showProgress("revenue", 80);
			revenueState.raw = rows;
			const { headers, checked, issues } = Validator.validateRevenue(rows);
			revenueState.headers = headers;
			revenueState.checked = checked;
			revenueState.issues = issues;
			updateFieldCheckboxes();
			updateBadges();
			renderTable(revenuePreview, revenueState, selectedFields);
			updateDataStats("revenue", checked.length, issues.length);
			showProgress("revenue", 100);
			setTimeout(() => hideProgress("revenue"), 500);
		}).catch((err) => {
			hideProgress("revenue");
			revenueUploadStatus.textContent = "读取失败";
			revenueValidBadge.textContent = "读取失败";
			revenueValidBadge.className = "badge warn";
			console.error("营收表读取失败：", err);
			// 更友好的错误提示
			const errorMsg = err.message || "未知错误";
			if (errorMsg.includes("CSV") || errorMsg.includes("xlsx")) {
				alert("文件格式错误：请确保上传的是有效的Excel文件(.xlsx/.xls)或CSV文件。");
			} else {
				alert("营收表读取失败：\n" + errorMsg + "\n\n请检查：\n1. 文件是否为xls/xlsx/csv格式\n2. 文件是否被其他程序占用\n3. 文件是否损坏");
			}
		});
	}

	function onRefundFile(file) {
		if (!file) return;
		refundUploadStatus.textContent = `已选择：${file.name}`;
		showProgress("refund", 20);
		const ext = getExt(file.name);
		ensureLibs().then(() => {
			showProgress("refund", 50);
			if (!(window && window.XLSX)) {
				if (ext === "csv") {
					return readFileSmart(file).then((u8) => {
						if (looksLikeXlsx(u8)) {
							alert("检测到该CSV文件实际是Excel工作簿（xlsx）。当前未加载Excel解析库，无法直接解析，请将文件正确保存为CSV，或把 xlsx.full.min.js 放入 vendor/ 目录后重试。");
							throw new Error("CSV masqueraded XLSX without library");
						}
						return decodeTextSmart(u8);
					}).then(parseCsv);
				}
				alert("当前环境未加载 Excel 解析库，无法直接解析 xlsx/xls。请将该文件另存为 CSV 后上传，或将 xlsx.full.min.js 放入 vendor/ 目录。");
				throw new Error("缺少XLSX库且文件并非CSV，无法解析。");
			}
			return readFileSmart(file).then(parseWorkbook);
		}).then((rows) => {
			showProgress("refund", 80);
			refundState.raw = rows;
			const { headers, checked, issues } = Validator.validateRefund(rows);
			refundState.headers = headers;
			refundState.checked = checked;
			refundState.issues = issues;
			updateFieldCheckboxes();
			updateBadges();
			renderTable(refundPreview, refundState, selectedFields);
			updateDataStats("refund", checked.length, issues.length);
			showProgress("refund", 100);
			setTimeout(() => hideProgress("refund"), 500);
		}).catch((err) => {
			hideProgress("refund");
			refundUploadStatus.textContent = "读取失败";
			refundValidBadge.textContent = "读取失败";
			refundValidBadge.className = "badge warn";
			console.error("退费表读取失败：", err);
			// 更友好的错误提示
			const errorMsg = err.message || "未知错误";
			if (errorMsg.includes("CSV") || errorMsg.includes("xlsx")) {
				alert("文件格式错误：请确保上传的是有效的Excel文件(.xlsx/.xls)或CSV文件。");
			} else {
				alert("退费表读取失败：\n" + errorMsg + "\n\n请检查：\n1. 文件是否为xls/xlsx/csv格式\n2. 文件是否被其他程序占用\n3. 文件是否损坏");
			}
		});
	}

	// 拖拽上传处理
	function setupDragAndDrop() {
		const revenueDropZone = document.getElementById("revenueDropZone");
		const refundDropZone = document.getElementById("refundDropZone");

		// 营收表拖拽
		if (revenueDropZone) {
			revenueDropZone.addEventListener("dragover", (e) => {
				e.preventDefault();
				revenueDropZone.classList.add("drag-over");
			});
			revenueDropZone.addEventListener("dragleave", () => {
				revenueDropZone.classList.remove("drag-over");
			});
			revenueDropZone.addEventListener("drop", (e) => {
				e.preventDefault();
				revenueDropZone.classList.remove("drag-over");
				const files = e.dataTransfer.files;
				if (files.length > 0) {
					onRevenueFile(files[0]);
				}
			});
		}

		// 退费表拖拽
		if (refundDropZone) {
			refundDropZone.addEventListener("dragover", (e) => {
				e.preventDefault();
				refundDropZone.classList.add("drag-over");
			});
			refundDropZone.addEventListener("dragleave", () => {
				refundDropZone.classList.remove("drag-over");
			});
			refundDropZone.addEventListener("drop", (e) => {
				e.preventDefault();
				refundDropZone.classList.remove("drag-over");
				const files = e.dataTransfer.files;
				if (files.length > 0) {
					onRefundFile(files[0]);
				}
			});
		}
	}

	// 确保DOM加载完成后再绑定事件
	function initEventListeners() {
		// 先获取DOM元素
		getDOMElements();
		
		console.log("初始化事件监听器...", {
			revenueInput: !!revenueInput,
			refundInput: !!refundInput,
			revenuePreview: !!revenuePreview,
			refundPreview: !!refundPreview
		});
		
		// 文件上传事件
		if (revenueInput) {
			revenueInput.addEventListener("change", (e) => {
				console.log("营收文件选择:", e.target.files);
				if (e.target.files && e.target.files[0]) {
					onRevenueFile(e.target.files[0]);
				} else {
					console.warn("未选择文件");
				}
			});
		} else {
			console.error("revenueInput元素未找到，请检查HTML中是否有id='revenueFile'的input元素");
		}
		
		if (refundInput) {
			refundInput.addEventListener("change", (e) => {
				console.log("退费文件选择:", e.target.files);
				if (e.target.files && e.target.files[0]) {
					onRefundFile(e.target.files[0]);
				} else {
					console.warn("未选择文件");
				}
			});
		} else {
			console.error("refundInput元素未找到，请检查HTML中是否有id='refundFile'的input元素");
		}
		
		// 仅显示有问题
		if (onlyIssues) {
			onlyIssues.addEventListener("change", () => {
				renderTable(revenuePreview, revenueState, selectedFields);
				renderTable(refundPreview, refundState, selectedFields);
			});
		}
		
		// 初始化拖拽上传
		setupDragAndDrop();
		
		// 字段选择按钮
		if (selectAllFields) {
			selectAllFields.addEventListener("click", () => {
				if (fieldCheckboxes) {
					fieldCheckboxes.querySelectorAll('input[type="checkbox"]').forEach(cb => {
						cb.checked = true;
						selectedFields.add(cb.value);
					});
					renderTable(revenuePreview, revenueState, selectedFields);
					renderTable(refundPreview, refundState, selectedFields);
				}
			});
		}
		
		if (deselectAllFields) {
			deselectAllFields.addEventListener("click", () => {
				if (fieldCheckboxes) {
					fieldCheckboxes.querySelectorAll('input[type="checkbox"]').forEach(cb => {
						cb.checked = false;
					});
					selectedFields.clear();
					renderTable(revenuePreview, revenueState, selectedFields);
					renderTable(refundPreview, refundState, selectedFields);
				}
			});
		}
		
		// 导出问题数据（只导出有问题的行）
		if (exportProblemsBtn) {
			exportProblemsBtn.addEventListener("click", () => {
				if (revenueState.issues.length === 0 && refundState.issues.length === 0) {
					alert("当前无问题数据可导出");
					return;
				}
				Validator.exportProblemData("问题数据导出.xlsx", revenueState.issues, refundState.issues);
			});
		}
		
		// 导出全部正确数据（只导出没有问题的数据）
		if (exportIssuesBtn) {
			exportIssuesBtn.addEventListener("click", () => {
				// 只导出没有问题的数据
				const allData = [
					...revenueState.checked.filter(x => !x.hasIssue).map(x => x.row),
					...refundState.checked.filter(x => !x.hasIssue).map(x => x.row),
				];
				if (!allData.length) {
					alert("当前无正确数据可导出，请先修复数据问题");
					return;
				}
				// 优先使用Excel导出，如果库不可用则自动使用CSV导出
				Validator.exportAllData("全部正确数据.xlsx", allData);
			});
		}
		
		// 进入分析页面
		if (goAnalysisBtn) {
			goAnalysisBtn.addEventListener("click", () => {
				// 两表均无问题则允许进入分析
				const okRevenue = revenueState.checked.length > 0 && revenueState.issues.length === 0;
				const okRefund = refundState.checked.length > 0 && refundState.issues.length === 0;
				if (!okRevenue || !okRefund) {
					alert("请先确保营收表与退费表均显示\"数据可用\"");
					return;
				}
				// 存储数据到 localStorage
				localStorage.setItem("revenueData", JSON.stringify(revenueState.checked.map(x => x.row)));
				localStorage.setItem("refundData", JSON.stringify(refundState.checked.map(x => x.row)));
				location.href = "analysis.html";
			});
		}
		
		console.log("事件监听器初始化完成");
	}
	
	// 在DOM加载完成后初始化，并确保库加载完成
	if (document.readyState === "loading") {
		document.addEventListener("DOMContentLoaded", () => {
			// 先确保库加载完成，再初始化事件
			ensureLibs().then(() => {
				console.log("库加载完成，初始化事件监听器");
				initEventListeners();
			}).catch((err) => {
				console.warn("库加载失败，但仍初始化事件监听器（CSV功能可用）:", err);
				initEventListeners();
			});
		});
	} else {
		// DOM已加载，立即初始化
		ensureLibs().then(() => {
			console.log("库加载完成，初始化事件监听器");
			setTimeout(initEventListeners, 0);
		}).catch((err) => {
			console.warn("库加载失败，但仍初始化事件监听器（CSV功能可用）:", err);
			setTimeout(initEventListeners, 0);
		});
	}
	
	// 切换编辑模式（这些按钮在initEventListeners中绑定）
	const toggleEditModeBtn = document.getElementById("toggleEditMode");
	const confirmEditBtn = document.getElementById("confirmEdit");
	const revalidateDataBtn = document.getElementById("revalidateData");
	
	function updateConfirmButtonState() {
		const hasEdits = editedData.revenue.size > 0 || editedData.refund.size > 0;
		if (confirmEditBtn) {
			confirmEditBtn.style.display = hasEdits ? "inline-block" : "none";
		}
	}
	
	toggleEditModeBtn?.addEventListener("click", () => {
		isEditMode = !isEditMode;
		toggleEditModeBtn.textContent = isEditMode ? "取消编辑" : "启用编辑";
		confirmEditBtn.style.display = isEditMode ? "inline-block" : "none";
		
		// 重新渲染表格
		renderTable(revenuePreview, revenueState, selectedFields);
		renderTable(refundPreview, refundState, selectedFields);
	});
	
	// 确认修改
	confirmEditBtn?.addEventListener("click", () => {
		const editCount = editedData.revenue.size + editedData.refund.size;
		if (editCount === 0) {
			alert("没有需要保存的修改");
			return;
		}
		
		if (!confirm(`您已修改 ${editCount} 处数据，确定要保存这些修改吗？\n\n保存后需要重新进行数据核查。`)) {
			return;
		}
		
		// 应用编辑到数据
		editedData.revenue.forEach((value, key) => {
			const [rowIndex, field] = key.split('_');
			const rowIdx = parseInt(rowIndex);
			if (revenueState.checked[rowIdx]) {
				revenueState.checked[rowIdx].row[field] = value;
				revenueState.raw[rowIdx][field] = value;
			}
		});
		
		editedData.refund.forEach((value, key) => {
			const [rowIndex, field] = key.split('_');
			const rowIdx = parseInt(rowIndex);
			if (refundState.checked[rowIdx]) {
				refundState.checked[rowIdx].row[field] = value;
				refundState.raw[rowIdx][field] = value;
			}
		});
		
		// 清空编辑记录
		editedData.revenue.clear();
		editedData.refund.clear();
		
		// 退出编辑模式
		isEditMode = false;
		toggleEditModeBtn.textContent = "启用编辑";
		confirmEditBtn.style.display = "none";
		revalidateDataBtn.style.display = "inline-block";
		
		// 重新渲染表格
		renderTable(revenuePreview, revenueState, selectedFields);
		renderTable(refundPreview, refundState, selectedFields);
		
		alert("修改已保存！请点击\"再次数据核查\"按钮重新验证数据。");
	});
	
	// 再次数据核查
	revalidateDataBtn?.addEventListener("click", () => {
		if (!confirm("确定要重新进行数据核查吗？这将重新验证所有数据。")) {
			return;
		}
		
		// 重新验证数据
		const revenueChecked = Validator.validateRevenue(revenueState.raw);
		const refundChecked = Validator.validateRefund(refundState.raw);
		
		revenueState.checked = revenueChecked.checked;
		revenueState.issues = revenueChecked.issues;
		
		refundState.checked = refundChecked.checked;
		refundState.issues = refundChecked.issues;
		
		// 更新统计
		updateDataStats("revenue", revenueState.checked.length, revenueState.issues.length);
		updateDataStats("refund", refundState.checked.length, refundState.issues.length);
		
		// 更新问题说明
		updateIssueSummary();
		
		// 更新徽章
		updateBadges();
		
		// 重新渲染表格
		renderTable(revenuePreview, revenueState, selectedFields);
		renderTable(refundPreview, refundState, selectedFields);
		
		// 隐藏再次核查按钮（如果没问题）
		if (revenueState.issues.length === 0 && refundState.issues.length === 0) {
			revalidateDataBtn.style.display = "none";
			alert("数据核查完成！所有数据均正确，可以导出全部正确数据。");
		} else {
			alert(`数据核查完成！发现 ${revenueState.issues.length + refundState.issues.length} 个问题，请继续修复。`);
		}
	});
	goAnalysisBtn.addEventListener("click", () => {
		// 两表均无问题则允许进入分析
		const okRevenue = revenueState.checked.length > 0 && revenueState.issues.length === 0;
		const okRefund = refundState.checked.length > 0 && refundState.issues.length === 0;
		if (!okRevenue || !okRefund) {
			alert("请先确保营收表与退费表均显示“数据可用”");
			return;
		}
		// 存储数据到 localStorage
		localStorage.setItem("revenueData", JSON.stringify(revenueState.checked.map(x => x.row)));
		localStorage.setItem("refundData", JSON.stringify(refundState.checked.map(x => x.row)));
		location.href = "analysis.html";
	});
})();


