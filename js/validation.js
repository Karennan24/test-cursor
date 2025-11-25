// 校验与工具函数
(function (global) {
	"use strict";

	function toArray(sheetJson) {
		return Array.isArray(sheetJson) ? sheetJson : [];
	}

	function normalizeDateValue(val, fieldName) {
		if (val == null || val === "") return "";
		
		// 只对日期/时间相关字段做格式化，避免金额等数字字段被误判
		const isDateField = fieldName && (
			/日期|时间|Date|Time|年月日|创建时间|收款日期|订单创建|时间戳/i.test(fieldName)
		);
		
		if (!isDateField) {
			// 非日期字段直接返回原始值
			return String(val);
		}
		
		if (val instanceof Date) {
			const y = val.getFullYear();
			const m = String(val.getMonth() + 1).padStart(2, "0");
			const d = String(val.getDate()).padStart(2, "0");
			return `${y}-${m}-${d}`;
		}
		
		// 处理Excel日期序列或可解析字符串（仅对日期字段）
		if (typeof val === "number") {
			// 金额通常不会超过几万，日期序列号通常很大（如40000+对应2009年）
			// 如果数字小于1000，很可能是金额，不要当成日期
			if (val < 1000) {
				return String(val);
			}
			const date = (typeof window !== "undefined" && window.XLSX && XLSX.SSF) ? XLSX.SSF.parse_date_code(val) : null;
			if (date && date.y && date.m && date.d) {
				const m = String(date.m).padStart(2, "0");
				const d = String(date.d).padStart(2, "0");
				return `${date.y}-${m}-${d}`;
			}
		}
		
		// 尝试解析为日期字符串
		const parsed = new Date(val);
		if (!isNaN(parsed.getTime())) {
			const y = parsed.getFullYear();
			const m = String(parsed.getMonth() + 1).padStart(2, "0");
			const d = String(parsed.getDate()).padStart(2, "0");
			// 如果解析出来的年份不合理（比如小于1900或大于2100），可能是金额被误判
			if (y >= 1900 && y <= 2100) {
				return `${y}-${m}-${d}`;
			}
		}
		
		return String(val);
	}

	function normalizeRowDates(row) {
		const out = { ...row };
		Object.keys(out).forEach((k) => {
			out[k] = normalizeDateValue(out[k], k);
		});
		return out;
	}

	function hasEmpty(value) {
		return value === undefined || value === null || String(value).trim() === "";
	}

	function normalizeCategory(val) {
		const v = String(val ?? "").trim();
		if (!v) return "";
		// 标准化常见分类
		if (/新生/i.test(v)) return "新生";
		if (/老生/i.test(v)) return "老生";
		if (/VIP/i.test(v)) return "VIP";
		if (/短期班|短期/i.test(v)) return "短期班";
		return v; // 其他原样
	}

	function validateRevenue(rows) {
		if (!rows || rows.length === 0) {
			return { headers: [], checked: [], issues: [] };
		}
		const headers = Object.keys(rows[0] || {});
		const required = ["教师姓名", "创建人姓名", "学科", "学生类型"];
		const issues = [];
		
		// 先过滤完全空的行，并保留原始索引
		const validRows = rows
			.map((raw, originalIdx) => ({ raw, originalIdx }))
			.filter(({ raw }) => {
				const hasAnyContent = Object.values(raw).some(v => String(v).trim().length > 0);
				return hasAnyContent;
			});
		
		const checked = validRows.map(({ raw, originalIdx }, filteredIdx) => {
			const row = normalizeRowDates(raw);
			const rowIssues = [];
			// 使用原始索引+1（因为Excel行号从1开始）显示问题行号
			const displayRowNum = originalIdx + 1;
			
			// 必填
			required.forEach((f) => {
				if (hasEmpty(row[f])) {
					rowIssues.push(`营收表第${displayRowNum}行【${f}】为空`);
				}
			});
			// 冲突校验：学生类型 与 班型
			const studentType = normalizeCategory(row["学生类型"]);
			const clazz = normalizeCategory(row["班型"]);
			if (studentType === "老生" && ["新生", "VIP", "短期班"].includes(clazz)) {
				rowIssues.push(`营收表第${displayRowNum}行 学生类型=老生 与 班型=${clazz} 存在冲突`);
			}
			const hasIssue = rowIssues.length > 0;
			if (hasIssue) {
				issues.push({ index: filteredIdx, originalIndex: originalIdx, reasons: rowIssues, row });
			}
			return { row, hasIssue, reasons: rowIssues, headers };
		});
		return { headers, checked, issues };
	}

	function validateRefund(rows) {
		if (!rows || rows.length === 0) {
			return { headers: [], checked: [], issues: [] };
		}
		const headers = Object.keys(rows[0] || {});
		const required = ["授课老师", "课程老师", "科目", "是否新生"];
		const issues = [];
		
		// 先过滤完全空的行，并保留原始索引
		const validRows = rows
			.map((raw, originalIdx) => ({ raw, originalIdx }))
			.filter(({ raw }) => {
				const hasAnyContent = Object.values(raw).some(v => String(v).trim().length > 0);
				return hasAnyContent;
			});
		
		const checked = validRows.map(({ raw, originalIdx }, filteredIdx) => {
			const row = normalizeRowDates(raw);
			const rowIssues = [];
			// 使用原始索引+1（因为Excel行号从1开始）显示问题行号
			const displayRowNum = originalIdx + 1;
			
			// 必填
			required.forEach((f) => {
				if (hasEmpty(row[f])) {
					rowIssues.push(`退费表第${displayRowNum}行【${f}】为空`);
				}
			});
			// 冲突校验：是否新生 与 班级（按与营收同样逻辑处理）
			const studentType = normalizeCategory(row["是否新生"]);
			const clazz = normalizeCategory(row["班级"]);
			if (studentType === "老生" && ["新生", "VIP", "短期班"].includes(clazz)) {
				rowIssues.push(`退费表第${displayRowNum}行 是否新生=老生 与 班级=${clazz} 存在冲突`);
			}
			const hasIssue = rowIssues.length > 0;
			if (hasIssue) {
				issues.push({ index: filteredIdx, originalIndex: originalIdx, reasons: rowIssues, row });
			}
			return { row, hasIssue, reasons: rowIssues, headers };
		});
		return { headers, checked, issues };
	}

	function toSheet(data) {
		return XLSX.utils.json_to_sheet(data);
	}

	function exportIssues(name, rows) {
		if (!window.XLSX || typeof saveAs === 'undefined') {
			// 使用CSV导出作为备用方案
			exportToCSV(rows, name);
			return;
		}
		
		const ws = toSheet(rows);
		const wb = XLSX.utils.book_new();
		XLSX.utils.book_append_sheet(wb, ws, "问题数据");
		const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
		saveAs(new Blob([wbout], { type: "application/octet-stream" }), name);
	}

	// CSV导出（不依赖外部库）
	function exportToCSV(data, filename) {
		if (!data || data.length === 0) {
			alert("没有数据可导出");
			return;
		}
		
		// 获取表头
		const headers = Object.keys(data[0]);
		
		// 构建CSV内容
		let csvContent = headers.map(h => escapeCSVField(h)).join(',') + '\n';
		
		data.forEach(row => {
			const values = headers.map(h => {
				const val = row[h];
				if (val == null) return '';
				return escapeCSVField(String(val));
			});
			csvContent += values.join(',') + '\n';
		});
		
		// 添加BOM以支持Excel正确识别UTF-8
		const BOM = '\uFEFF';
		const blob = new Blob([BOM + csvContent], { type: 'text/csv;charset=utf-8;' });
		
		// 使用a标签下载（不依赖FileSaver）
		const link = document.createElement('a');
		const url = URL.createObjectURL(blob);
		link.href = url;
		link.download = filename.replace('.xlsx', '.csv').replace('.xls', '.csv');
		document.body.appendChild(link);
		link.click();
		document.body.removeChild(link);
		URL.revokeObjectURL(url);
	}
	
	function escapeCSVField(field) {
		if (field == null) return '';
		const str = String(field);
		// 如果包含逗号、引号或换行符，需要用引号包裹
		if (str.includes(',') || str.includes('"') || str.includes('\n') || str.includes('\r')) {
			return '"' + str.replace(/"/g, '""') + '"';
		}
		return str;
	}

	function exportAllData(name, allData) {
		// 优先使用Excel导出，如果库不可用则使用CSV
		if (!window.XLSX || typeof saveAs === 'undefined') {
			// 使用CSV导出作为备用方案
			exportToCSV(allData, name);
			return;
		}
		
		// 准备数据：移除内部标记字段，但保留用于标记
		const cleanData = allData.map(row => {
			const clean = { ...row };
			delete clean._source;
			delete clean._hasIssue;
			delete clean._reasons;
			return clean;
		});
		
		const ws = XLSX.utils.json_to_sheet(cleanData);
		
		// 获取工作表范围
		const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
		
		// 设置表头样式
		for (let col = range.s.c; col <= range.e.c; col++) {
			const cellAddr = XLSX.utils.encode_cell({ r: 0, c: col });
			if (!ws[cellAddr]) continue;
			ws[cellAddr].s = {
				fill: { fgColor: { rgb: "E5E7EB" } },
				font: { bold: true },
				border: {
					top: { style: "thin" },
					bottom: { style: "thin" },
					left: { style: "thin" },
					right: { style: "thin" }
				}
			};
		}
		
		// 标记有问题数据行（橘黄色背景）
		allData.forEach((row, idx) => {
			if (row._hasIssue) {
				for (let col = range.s.c; col <= range.e.c; col++) {
					const cellAddr = XLSX.utils.encode_cell({ r: idx + 1, c: col });
					if (!ws[cellAddr]) continue;
					if (!ws[cellAddr].s) ws[cellAddr].s = {};
					ws[cellAddr].s.fill = { fgColor: { rgb: "FED7AA" } }; // 橘黄色背景
				}
			}
		});
		
		// 设置列宽
		const colWidths = [];
		const headers = Object.keys(cleanData[0] || {});
		headers.forEach((h, idx) => {
			const maxLength = Math.max(
				h.length,
				...cleanData.slice(0, 100).map(r => String(r[h] || '').length)
			);
			colWidths.push({ wch: Math.min(maxLength + 2, 50) });
		});
		ws['!cols'] = colWidths;
		
		const wb = XLSX.utils.book_new();
		XLSX.utils.book_append_sheet(wb, ws, "全部数据");
		
		const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
		saveAs(new Blob([wbout], { type: "application/octet-stream" }), name);
	}
	
	// 导出问题数据（只导出有问题的行）
	function exportProblemData(name, revenueIssues, refundIssues) {
		const problemData = [
			...revenueIssues.map(issue => ({ ...issue.row, _source: '营收表', _problem: issue.reasons.join('；') })),
			...refundIssues.map(issue => ({ ...issue.row, _source: '退费表', _problem: issue.reasons.join('；') })),
		];
		
		if (problemData.length === 0) {
			alert("当前无问题数据可导出");
			return;
		}
		
		// 优先使用Excel导出，如果库不可用则使用CSV
		if (!window.XLSX || typeof saveAs === 'undefined') {
			exportToCSV(problemData, name);
			return;
		}
		
		// 准备数据：移除内部标记字段
		const cleanData = problemData.map(row => {
			const clean = { ...row };
			delete clean._source;
			delete clean._problem;
			// 添加问题说明列
			clean['问题说明'] = row._problem;
			clean['数据来源'] = row._source;
			return clean;
		});
		
		const ws = XLSX.utils.json_to_sheet(cleanData);
		
		// 获取工作表范围
		const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
		
		// 设置表头样式
		for (let col = range.s.c; col <= range.e.c; col++) {
			const cellAddr = XLSX.utils.encode_cell({ r: 0, c: col });
			if (!ws[cellAddr]) continue;
			ws[cellAddr].s = {
				fill: { fgColor: { rgb: "E5E7EB" } },
				font: { bold: true },
				border: {
					top: { style: "thin" },
					bottom: { style: "thin" },
					left: { style: "thin" },
					right: { style: "thin" }
				}
			};
		}
		
		// 标记所有问题数据行（橘黄色背景）
		for (let rowIdx = 1; rowIdx <= cleanData.length; rowIdx++) {
			for (let col = range.s.c; col <= range.e.c; col++) {
				const cellAddr = XLSX.utils.encode_cell({ r: rowIdx, c: col });
				if (!ws[cellAddr]) continue;
				if (!ws[cellAddr].s) ws[cellAddr].s = {};
				ws[cellAddr].s.fill = { fgColor: { rgb: "FED7AA" } }; // 橘黄色背景
			}
		}
		
		// 设置列宽
		const colWidths = [];
		const headers = Object.keys(cleanData[0] || {});
		headers.forEach((h, idx) => {
			const maxLength = Math.max(
				h.length,
				...cleanData.slice(0, 100).map(r => String(r[h] || '').length)
			);
			colWidths.push({ wch: Math.min(maxLength + 2, 50) });
		});
		ws['!cols'] = colWidths;
		
		const wb = XLSX.utils.book_new();
		XLSX.utils.book_append_sheet(wb, ws, "问题数据");
		
		const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
		saveAs(new Blob([wbout], { type: "application/octet-stream" }), name);
	}

	global.Validator = {
		toArray,
		validateRevenue,
		validateRefund,
		exportIssues,
		exportAllData,
		exportProblemData,
	};
})(window);


