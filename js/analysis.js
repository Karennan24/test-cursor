(function () {
	"use strict";
	
	let currentSalesData = [];
	let selectedSalesNames = new Set(); // 改为Set支持多选
	let pendingSelections = new Set(); // 待应用的选择（多选模式下累积）
	let isMultiSelectMode = false; // 是否开启多选模式
	let chartInstances = { bar: null, pie: null };
	let animationFrameId = null;

	// 执行状态标志，防止重复执行
	let isFiltering = false;
	let pendingFilterCancel = false;
	
	// 优化的防抖函数，支持取消
	function debounce(func, wait) {
		let timeout;
		let lastArgs;
		const debounced = function executedFunction(...args) {
			lastArgs = args;
			clearTimeout(timeout);
			timeout = setTimeout(() => {
				if (!pendingFilterCancel) {
					func(...args);
				}
				pendingFilterCancel = false;
			}, wait);
		};
		debounced.cancel = () => {
			clearTimeout(timeout);
			pendingFilterCancel = true;
		};
		debounced.flush = () => {
			clearTimeout(timeout);
			if (lastArgs) {
				func(...lastArgs);
			}
		};
		return debounced;
	}

	function getData() {
		try {
			const revenue = JSON.parse(localStorage.getItem("revenueData") || "[]");
			const refund = JSON.parse(localStorage.getItem("refundData") || "[]");
			return { revenue, refund };
		} catch {
			return { revenue: [], refund: [] };
		}
	}

	function sum(arr) {
		return arr.reduce((a, b) => a + (Number(b) || 0), 0);
	}

	function formatToWanYuan(amount) {
		if (amount === 0) return "0万元";
		const wan = amount / 10000;
		if (wan >= 1000) {
			return `${(wan / 1000).toFixed(2)}千万元`;
		}
		return `${wan.toFixed(2)}万元`;
	}

	function renderSummary({ revenue, refund }) {
		const totalRevenue = sum(revenue.map(r => Number(r["课程拆分金额"]) || 0));
		const totalRefund = sum(refund.map(r => Number(r["退费金额"]) || 0));
		const net = totalRevenue - totalRefund;
		const refundRate = totalRevenue > 0 ? (totalRefund / totalRevenue * 100) : 0;
		const netRevenueRate = totalRevenue > 0 ? (net / totalRevenue * 100) : 0;
		
		// 使用新的卡片式布局
		const el = document.getElementById("summaryCards");
		el.innerHTML = `
			<div class="metric-card ${refundRate < 5 ? 'success' : refundRate < 10 ? 'warning' : 'danger'}">
				<div class="metric-label">总营收</div>
				<div class="metric-value">${formatToWanYuan(totalRevenue)}</div>
				<div class="metric-trend">${revenue.length}笔订单</div>
			</div>
			<div class="metric-card ${refundRate < 5 ? 'success' : refundRate < 10 ? 'warning' : 'danger'}">
				<div class="metric-label">总退费</div>
				<div class="metric-value">${formatToWanYuan(totalRefund)}</div>
				<div class="metric-trend">退费率 ${refundRate.toFixed(1)}%</div>
			</div>
			<div class="metric-card ${netRevenueRate > 70 ? 'success' : netRevenueRate > 50 ? 'warning' : 'danger'}">
				<div class="metric-label">净营收</div>
				<div class="metric-value">${formatToWanYuan(net)}</div>
				<div class="metric-trend">净营收率 ${netRevenueRate.toFixed(1)}%</div>
			</div>
		`;
	}

	function aggregateBySales({ revenue, refund }) {
		// 字段映射：
		// 营收表：创建人姓名 => 销售人员；课程拆分金额 => 营收
		// 退费表：课程老师   => 销售人员；退费金额   => 退费金额
		const salesMap = new Map();
		function addAgg(name, rev, ref) {
			if (!name) return;
			const cur = salesMap.get(name) || { sales: name, revenue: 0, refund: 0 };
			cur.revenue += (Number(rev) || 0);
			cur.refund += (Number(ref) || 0);
			salesMap.set(name, cur);
		}
		revenue.forEach(r => addAgg(r["创建人姓名"], r["课程拆分金额"], 0));
		refund.forEach(r => addAgg(r["课程老师"], 0, r["退费金额"]));
		return Array.from(salesMap.values()).map(x => ({
			销售人员: x.sales,
			营收: Number(x.revenue.toFixed(2)),
			退费金额: Number(x.refund.toFixed(2)),
			净营收: Number((x.revenue - x.refund).toFixed(2)),
		}));
	}

	function renderTable(containerId, rows, filterNames = null) {
		const container = document.getElementById(containerId);
		container.innerHTML = "";
		if (!rows.length) {
			container.textContent = "暂无数据";
			return;
		}
		
		// 如果指定了筛选，只显示选中的销售人员
		const displayRows = filterNames && filterNames.size > 0
			? rows.filter(r => filterNames.has(r["销售人员"]))
			: rows;
		
		const headers = Object.keys(rows[0]);
		const table = document.createElement("table");
		const thead = document.createElement("thead");
		const trh = document.createElement("tr");
		headers.forEach(h => {
			const th = document.createElement("th");
			th.textContent = h;
			trh.appendChild(th);
		});
		thead.appendChild(trh);
		table.appendChild(thead);

		const tbody = document.createElement("tbody");
		displayRows.forEach(row => {
			const tr = document.createElement("tr");
			tr.dataset.sales = row["销售人员"];
			tr.classList.add("sales-row");
			if (selectedSalesNames.has(row["销售人员"])) {
				tr.classList.add("selected");
			}
			// 多选模式下，显示待应用的选择状态
			if (isMultiSelectMode && pendingSelections.has(row["销售人员"])) {
				tr.classList.add("pending");
			}
			headers.forEach((h, idx) => {
				const td = document.createElement("td");
				if (typeof row[h] === 'number') {
					td.textContent = row[h].toFixed(2);
				} else {
					td.textContent = row[h];
					// 销售人员列添加点击事件（支持多选）
					if (h === "销售人员") {
						td.style.cursor = "pointer";
						td.style.color = "var(--brand)";
						td.style.textDecoration = "underline";
						td.addEventListener("click", (e) => {
							e.stopPropagation();
							toggleSales(row["销售人员"], rows);
						});
					}
				}
				tr.appendChild(td);
			});
			tbody.appendChild(tr);
		});
		
		// 添加总计行（使用原始rows计算，不是filtered）
		const totalRow = document.createElement("tr");
		totalRow.className = "total-row";
		headers.forEach(h => {
			const td = document.createElement("td");
			if (h === "销售人员") {
				if (filterNames && filterNames.size > 0) {
					const names = Array.from(filterNames);
					td.textContent = names.length === 1 ? `${names[0]} - 合计` : `已选${names.length}人 - 合计`;
				} else {
					td.textContent = "合计";
				}
			} else {
				const total = (filterNames && filterNames.size > 0 ? displayRows : rows).reduce((sum, row) => sum + (Number(row[h]) || 0), 0);
				td.textContent = total.toFixed(2);
			}
			totalRow.appendChild(td);
		});
		tbody.appendChild(totalRow);
		
		table.appendChild(tbody);
		container.appendChild(table);
		return rows; // 返回完整数据用于图表
	}
	
	function toggleSales(salesName, allRows) {
		if (isMultiSelectMode) {
			// 多选模式：累积选择，不立即更新
			if (pendingSelections.has(salesName)) {
				pendingSelections.delete(salesName);
			} else {
				pendingSelections.add(salesName);
			}
			updatePendingInfo();
		} else {
			// 单选模式：立即更新
			if (selectedSalesNames.has(salesName)) {
				selectedSalesNames.clear();
			} else {
				selectedSalesNames.clear();
				selectedSalesNames.add(salesName);
			}
			applySelections(allRows);
		}
	}
	
	function applySelections(allRows) {
		// 使用requestAnimationFrame优化性能
		if (animationFrameId) {
			cancelAnimationFrame(animationFrameId);
		}
		animationFrameId = requestAnimationFrame(() => {
			// 更新表格
			renderTable("salesAgg", allRows, selectedSalesNames.size > 0 ? selectedSalesNames : null);
			
			// 高亮图表
			highlightChartElement(selectedSalesNames);
			
			// 显示筛选信息
			updateFilterInfo(allRows);
		});
	}
	
	function applyPendingSelections(allRows) {
		// 将待应用的选择应用到实际筛选
		selectedSalesNames = new Set(pendingSelections);
		applySelections(allRows);
	}
	
	function clearSelections(allRows) {
		selectedSalesNames.clear();
		pendingSelections.clear();
		applySelections(allRows);
	}
	
	function updatePendingInfo() {
		const filterInfo = document.getElementById("filterInfo");
		const applyBtn = document.getElementById("applyFilter");
		if (pendingSelections.size > 0) {
			if (applyBtn) {
				applyBtn.style.display = "inline-block";
				applyBtn.textContent = `应用筛选 (${pendingSelections.size})`;
			}
			if (filterInfo) {
				const names = Array.from(pendingSelections);
				filterInfo.textContent = `待应用: ${names.join("、")}`;
				filterInfo.style.color = "var(--warn)";
			}
		} else {
			if (applyBtn) applyBtn.style.display = "none";
			if (filterInfo) {
				filterInfo.textContent = "";
			}
		}
	}
	
	function updateFilterInfo(allRows) {
		const resetBtn = document.getElementById("resetFilter");
		const filterInfo = document.getElementById("filterInfo");
		if (selectedSalesNames.size > 0) {
			if (resetBtn) {
				resetBtn.style.display = "inline-block";
				resetBtn.onclick = () => {
					clearSelections(allRows);
				};
			}
			if (filterInfo) {
				const names = Array.from(selectedSalesNames);
				filterInfo.textContent = `已选${names.length}人: ${names.join("、")}`;
				filterInfo.style.color = "var(--muted)";
			}
		} else {
			if (resetBtn) resetBtn.style.display = "none";
			if (filterInfo) {
				filterInfo.textContent = "";
			}
		}
	}
	
	function highlightChartElement(selectedNames) {
		// 重绘柱状图，高亮对应销售人员
		if (chartInstances.bar && currentSalesData.length > 0) {
			drawBarChart("revenueRefundChart", currentSalesData, selectedNames);
		}
		// 重绘饼图，高亮对应销售人员
		if (chartInstances.pie && currentSalesData.length > 0) {
			drawPieChart("netRevenueChart", currentSalesData, selectedNames);
		}
	}
	
	function resetChartHighlight() {
		if (chartInstances.bar && currentSalesData.length > 0) {
			drawBarChart("revenueRefundChart", currentSalesData, new Set());
		}
		if (chartInstances.pie && currentSalesData.length > 0) {
			drawPieChart("netRevenueChart", currentSalesData, new Set());
		}
	}

	function drawBarChart(canvasId, salesData, highlightNames = new Set()) {
		const canvas = document.getElementById(canvasId);
		if (!canvas) return;
		const ctx = canvas.getContext("2d");
		// 清除画布
		ctx.clearRect(0, 0, canvas.width, canvas.height);
		const tip = document.getElementById("chartTip");
		
		const width = canvas.width;
		const height = canvas.height;
		const padding = { top: 40, right: 40, bottom: 60, left: 80 };
		const chartWidth = width - padding.left - padding.right;
		const chartHeight = height - padding.top - padding.bottom;
		
		// 清空画布
		ctx.clearRect(0, 0, width, height);
		
		if (!salesData || salesData.length === 0) {
			ctx.fillStyle = "#9ca3af";
			ctx.font = "16px sans-serif";
			ctx.textAlign = "center";
			ctx.fillText("暂无数据", width / 2, height / 2);
			return;
		}
		
		// 计算最大值
		const maxValue = Math.max(...salesData.map(s => Math.max(s.营收, s.退费金额)));
		const maxY = Math.ceil(maxValue * 1.1);
		
		const barWidth = chartWidth / (salesData.length * 2.5);
		const gap = barWidth * 0.3;
		
		// 绘制坐标轴
		ctx.strokeStyle = "#4b5563";
		ctx.lineWidth = 1;
		ctx.beginPath();
		ctx.moveTo(padding.left, padding.top);
		ctx.lineTo(padding.left, padding.top + chartHeight);
		ctx.lineTo(padding.left + chartWidth, padding.top + chartHeight);
		ctx.stroke();
		
		// 绘制Y轴刻度
		ctx.fillStyle = "#9ca3af";
		ctx.font = "11px sans-serif";
		ctx.textAlign = "right";
		for (let i = 0; i <= 5; i++) {
			const value = (maxY / 5) * i;
			const y = padding.top + chartHeight - (value / maxY) * chartHeight;
			ctx.fillText(value.toFixed(0), padding.left - 10, y + 4);
		}
		
		// 绘制柱状图（科技感增强）
		salesData.forEach((sales, idx) => {
			const x = padding.left + idx * barWidth * 2.5 + gap;
			const isHighlighted = highlightNames.has(sales.销售人员);
			const highlightOffset = isHighlighted ? 8 : 0;
			const glowSize = isHighlighted ? 12 : 0;
			
			// 营收柱 - 使用渐变和阴影
			const revenueHeight = (sales.营收 / maxY) * chartHeight;
			const revenueGradient = ctx.createLinearGradient(x, padding.top + chartHeight - revenueHeight, x, padding.top + chartHeight);
			revenueGradient.addColorStop(0, isHighlighted ? "#60a5fa" : "#3b82f6");
			revenueGradient.addColorStop(1, isHighlighted ? "#2563eb" : "#1e40af");
			ctx.fillStyle = revenueGradient;
			ctx.shadowBlur = isHighlighted ? glowSize : 0;
			ctx.shadowColor = isHighlighted ? "#60a5fa" : "transparent";
			ctx.fillRect(x - highlightOffset, padding.top + chartHeight - revenueHeight - highlightOffset, barWidth + highlightOffset * 2, revenueHeight + highlightOffset);
			if (isHighlighted) {
				ctx.strokeStyle = "#fbbf24";
				ctx.lineWidth = 3;
				ctx.shadowBlur = 0;
				ctx.strokeRect(x - highlightOffset, padding.top + chartHeight - revenueHeight - highlightOffset, barWidth + highlightOffset * 2, revenueHeight + highlightOffset);
			}
			
			// 退费柱 - 使用渐变和阴影
			const refundHeight = (sales.退费金额 / maxY) * chartHeight;
			const refundGradient = ctx.createLinearGradient(x + barWidth + gap, padding.top + chartHeight - refundHeight, x + barWidth + gap, padding.top + chartHeight);
			refundGradient.addColorStop(0, isHighlighted ? "#f87171" : "#ef4444");
			refundGradient.addColorStop(1, isHighlighted ? "#dc2626" : "#b91c1c");
			ctx.fillStyle = refundGradient;
			ctx.shadowBlur = isHighlighted ? glowSize : 0;
			ctx.shadowColor = isHighlighted ? "#f87171" : "transparent";
			ctx.fillRect(x + barWidth + gap - highlightOffset, padding.top + chartHeight - refundHeight - highlightOffset, barWidth + highlightOffset * 2, refundHeight + highlightOffset);
			if (isHighlighted) {
				ctx.strokeStyle = "#fbbf24";
				ctx.lineWidth = 3;
				ctx.shadowBlur = 0;
				ctx.strokeRect(x + barWidth + gap - highlightOffset, padding.top + chartHeight - refundHeight - highlightOffset, barWidth + highlightOffset * 2, refundHeight + highlightOffset);
			}
			
			// 重置阴影
			ctx.shadowBlur = 0;
			
			// X轴标签（截断长名字）
			ctx.fillStyle = "#e5e7eb";
			ctx.font = "10px sans-serif";
			ctx.textAlign = "center";
			const label = sales.销售人员.length > 4 ? sales.销售人员.substring(0, 4) + "..." : sales.销售人员;
			ctx.save();
			ctx.translate(x + barWidth, padding.top + chartHeight + 15);
			ctx.rotate(-Math.PI / 4);
			ctx.fillText(label, 0, 0);
			ctx.restore();
		});
		
		// 图例
		ctx.fillStyle = "#3b82f6";
		ctx.fillRect(padding.left + chartWidth - 100, padding.top - 25, 12, 12);
		ctx.fillStyle = "#e5e7eb";
		ctx.font = "11px sans-serif";
		ctx.textAlign = "left";
		ctx.fillText("营收", padding.left + chartWidth - 85, padding.top - 15);
		
		ctx.fillStyle = "#ef4444";
		ctx.fillRect(padding.left + chartWidth - 50, padding.top - 25, 12, 12);
		ctx.fillStyle = "#e5e7eb";
		ctx.fillText("退费", padding.left + chartWidth - 35, padding.top - 15);
		
		// 鼠标交互
		let hoveredIndex = -1;
		canvas.addEventListener("mousemove", (e) => {
			const rect = canvas.getBoundingClientRect();
			const x = e.clientX - rect.left;
			const y = e.clientY - rect.top;
			
			let found = false;
			salesData.forEach((sales, idx) => {
				const barX = padding.left + idx * barWidth * 2.5 + gap;
				if (x >= barX && x <= barX + barWidth * 2 + gap && y >= padding.top && y <= padding.top + chartHeight) {
					found = true;
					hoveredIndex = idx;
					canvas.style.cursor = "pointer";
					if (tip) {
						tip.style.display = "block";
						tip.style.left = e.pageX + 10 + "px";
						tip.style.top = e.pageY + 10 + "px";
						tip.innerHTML = `
							<div><strong>${sales.销售人员}</strong></div>
							<div>营收: ¥${sales.营收.toFixed(2)}</div>
							<div>退费: ¥${sales.退费金额.toFixed(2)}</div>
							<div>净营收: ¥${sales.净营收.toFixed(2)}</div>
						`;
					}
				}
			});
			if (!found) {
				hoveredIndex = -1;
				canvas.style.cursor = "default";
				if (tip) tip.style.display = "none";
			}
		});
		
		canvas.addEventListener("click", (e) => {
			if (hoveredIndex >= 0) {
				const sales = salesData[hoveredIndex];
				toggleSales(sales.销售人员, salesData);
			}
		});
		
		chartInstances.bar = { canvas, salesData, highlightNames };
	}
	
	function drawPieChart(canvasId, salesData, highlightNames = new Set()) {
		const canvas = document.getElementById(canvasId);
		if (!canvas) return;
		const ctx = canvas.getContext("2d");
		// 清除画布
		ctx.clearRect(0, 0, canvas.width, canvas.height);
		const tip = document.getElementById("pieChartTip");
		
		const width = canvas.width;
		const height = canvas.height;
		const centerX = width / 2;
		const centerY = height / 2;
		const radius = Math.min(width, height) / 2 - 60;
		
		ctx.clearRect(0, 0, width, height);
		
		if (!salesData || salesData.length === 0) {
			ctx.fillStyle = "#9ca3af";
			ctx.font = "16px sans-serif";
			ctx.textAlign = "center";
			ctx.fillText("暂无数据", width / 2, height / 2);
			return;
		}
		
		// 过滤掉净营收为负数的
		const positiveData = salesData.filter(s => s.净营收 > 0);
		const total = positiveData.reduce((sum, s) => sum + s.净营收, 0);
		
		if (total === 0) {
			ctx.fillStyle = "#9ca3af";
			ctx.font = "16px sans-serif";
			ctx.textAlign = "center";
			ctx.fillText("暂无正净营收数据", width / 2, height / 2);
			return;
		}
		
		// 颜色方案
		const colors = [
			"#3b82f6", "#10b981", "#f59e0b", "#ef4444", "#8b5cf6",
			"#06b6d4", "#ec4899", "#84cc16", "#f97316", "#6366f1"
		];
		
		let currentAngle = -Math.PI / 2;
		let hoveredIndex = -1;
		
		// 绘制饼图（科技感增强）
		positiveData.forEach((sales, idx) => {
			const angle = (sales.净营收 / total) * 2 * Math.PI;
			const isHovered = hoveredIndex === idx;
			const isHighlighted = highlightNames.has(sales.销售人员);
			const sliceRadius = radius + (isHovered ? 10 : 0) + (isHighlighted ? 15 : 0);
			
			// 创建渐变
			const gradient = ctx.createRadialGradient(
				centerX + Math.cos(currentAngle + angle / 2) * sliceRadius * 0.3,
				centerY + Math.sin(currentAngle + angle / 2) * sliceRadius * 0.3,
				0,
				centerX, centerY, sliceRadius
			);
			const baseColor = colors[idx % colors.length];
			gradient.addColorStop(0, isHighlighted ? lightenColor(baseColor, 30) : lightenColor(baseColor, 20));
			gradient.addColorStop(1, baseColor);
			
			ctx.beginPath();
			ctx.moveTo(centerX, centerY);
			ctx.arc(centerX, centerY, sliceRadius, currentAngle, currentAngle + angle);
			ctx.closePath();
			ctx.fillStyle = gradient;
			
			// 添加发光效果
			if (isHighlighted) {
				ctx.shadowBlur = 20;
				ctx.shadowColor = baseColor;
			}
			ctx.fill();
			ctx.shadowBlur = 0;
			
			ctx.strokeStyle = isHighlighted ? "#fbbf24" : "#1f2937";
			ctx.lineWidth = isHighlighted ? 4 : 2;
			ctx.stroke();
			
			// 绘制标签
			const labelAngle = currentAngle + angle / 2;
			const labelRadius = radius * 0.7;
			const labelX = centerX + Math.cos(labelAngle) * labelRadius;
			const labelY = centerY + Math.sin(labelAngle) * labelRadius;
			
			ctx.fillStyle = "#e5e7eb";
			ctx.font = "11px sans-serif";
			ctx.textAlign = "center";
			ctx.textBaseline = "middle";
			const percentage = ((sales.净营收 / total) * 100).toFixed(1);
			if (percentage >= 5) {
				ctx.fillText(`${percentage}%`, labelX, labelY);
			}
			
			currentAngle += angle;
		});
		
		// 图例
		let legendY = 30;
		positiveData.forEach((sales, idx) => {
			ctx.fillStyle = colors[idx % colors.length];
			ctx.fillRect(20, legendY, 12, 12);
			ctx.fillStyle = "#e5e7eb";
			ctx.font = "10px sans-serif";
			ctx.textAlign = "left";
			const label = sales.销售人员.length > 8 ? sales.销售人员.substring(0, 8) + "..." : sales.销售人员;
			ctx.fillText(label, 35, legendY + 9);
			legendY += 18;
		});
		
		// 鼠标交互
		canvas.addEventListener("mousemove", (e) => {
			const rect = canvas.getBoundingClientRect();
			const x = e.clientX - rect.left - centerX;
			const y = e.clientY - rect.top - centerY;
			const distance = Math.sqrt(x * x + y * y);
			
			if (distance <= radius + 10) {
				let angle = Math.atan2(y, x);
				if (angle < -Math.PI / 2) angle += 2 * Math.PI;
				angle += Math.PI / 2;
				if (angle >= 2 * Math.PI) angle -= 2 * Math.PI;
				
				let currentAngle = -Math.PI / 2;
				let found = false;
				positiveData.forEach((sales, idx) => {
					const sliceAngle = (sales.净营收 / total) * 2 * Math.PI;
					if (angle >= currentAngle && angle <= currentAngle + sliceAngle) {
						found = true;
						hoveredIndex = idx;
						canvas.style.cursor = "pointer";
						drawPieChart(canvasId, salesData); // 重绘以显示悬停效果
						if (tip) {
							tip.style.display = "block";
							tip.style.left = e.pageX + 10 + "px";
							tip.style.top = e.pageY + 10 + "px";
							const percentage = ((sales.净营收 / total) * 100).toFixed(1);
							tip.innerHTML = `
								<div><strong>${sales.销售人员}</strong></div>
								<div>净营收: ¥${sales.净营收.toFixed(2)}</div>
								<div>占比: ${percentage}%</div>
							`;
						}
					}
					currentAngle += sliceAngle;
				});
				if (!found) {
					hoveredIndex = -1;
					canvas.style.cursor = "default";
					if (tip) tip.style.display = "none";
				}
			} else {
				hoveredIndex = -1;
				canvas.style.cursor = "default";
				if (tip) tip.style.display = "none";
			}
		});
		
		canvas.addEventListener("click", (e) => {
			if (hoveredIndex >= 0) {
				const sales = positiveData[hoveredIndex];
				toggleSales(sales.销售人员, salesData);
			}
		});
		
		chartInstances.pie = { canvas, salesData, highlightNames };
	}
	

	// 检查数据是否为空
	function checkDataEmpty(data) {
		const isEmpty = !data.revenue || data.revenue.length === 0 || 
		                !data.refund || data.refund.length === 0;
		
		const emptyTip = document.getElementById("emptyDataTip");
		const filterSection = document.getElementById("filterSection");
		const tabsSection = document.getElementById("tabsSection");
		
		if (isEmpty) {
			if (emptyTip) emptyTip.style.display = "block";
			if (filterSection) filterSection.style.display = "none";
			if (tabsSection) tabsSection.style.display = "none";
			// 隐藏所有选项卡内容
			document.querySelectorAll(".tab-content").forEach(tab => {
				tab.style.display = "none";
			});
			return true;
		} else {
			if (emptyTip) emptyTip.style.display = "none";
			if (filterSection) filterSection.style.display = "block";
			if (tabsSection) tabsSection.style.display = "block";
			// 显示默认选项卡（概览）
			const overviewTab = document.getElementById("overviewTab");
			if (overviewTab) {
				overviewTab.style.display = "block";
				overviewTab.classList.add("active");
			}
			return false;
		}
	}

	// 时间范围筛选
	function filterByDateRange(data, startDate, endDate) {
		if (!startDate && !endDate) return data;
		
		const filteredRevenue = data.revenue.filter(row => {
			const dateStr = row["收款日期"] || row["订单创建时间"] || "";
			if (!dateStr) return true; // 没有日期的保留
			const rowDate = new Date(dateStr);
			if (isNaN(rowDate.getTime())) return true; // 无效日期保留
			
			if (startDate && rowDate < new Date(startDate)) return false;
			if (endDate && rowDate > new Date(endDate + "T23:59:59")) return false;
			return true;
		});
		
		const filteredRefund = data.refund.filter(row => {
			const dateStr = row["收款日期"] || row["申请退款日期"] || "";
			if (!dateStr) return true;
			const rowDate = new Date(dateStr);
			if (isNaN(rowDate.getTime())) return true;
			
			if (startDate && rowDate < new Date(startDate)) return false;
			if (endDate && rowDate > new Date(endDate + "T23:59:59")) return false;
			return true;
		});
		
		return { revenue: filteredRevenue, refund: filteredRefund };
	}

	// 导出分析结果
	function exportAnalysisResults(salesData) {
		if (!salesData || salesData.length === 0) {
			alert("没有数据可导出");
			return;
		}
		
		// 检查是否有导出库
		if (!window.XLSX || typeof saveAs === 'undefined') {
			// 使用CSV导出
			exportToCSV(salesData);
			return;
		}
		
		// Excel导出
		const ws = XLSX.utils.json_to_sheet(salesData);
		const wb = XLSX.utils.book_new();
		XLSX.utils.book_append_sheet(wb, ws, "销售分析");
		
		// 设置列宽
		const colWidths = [
			{ wch: 15 }, // 销售人员
			{ wch: 15 }, // 营收
			{ wch: 15 }, // 退费金额
			{ wch: 15 }  // 净营收
		];
		ws['!cols'] = colWidths;
		
		const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
		const fileName = `销售分析_${new Date().toISOString().split('T')[0]}.xlsx`;
		saveAs(new Blob([wbout], { type: "application/octet-stream" }), fileName);
	}

	// CSV导出
	function exportToCSV(data) {
		if (!data || data.length === 0) return;
		
		const headers = Object.keys(data[0]);
		let csv = headers.map(h => `"${h}"`).join(',') + '\n';
		
		data.forEach(row => {
			const values = headers.map(h => {
				const val = row[h];
				return `"${val != null ? String(val).replace(/"/g, '""') : ''}"`;
			});
			csv += values.join(',') + '\n';
		});
		
		const BOM = '\uFEFF';
		const blob = new Blob([BOM + csv], { type: 'text/csv;charset=utf-8;' });
		const link = document.createElement('a');
		const url = URL.createObjectURL(blob);
		link.href = url;
		link.download = `销售分析_${new Date().toISOString().split('T')[0]}.csv`;
		document.body.appendChild(link);
		link.click();
		document.body.removeChild(link);
		URL.revokeObjectURL(url);
	}

	let originalData = null; // 保存原始数据用于重置

	function main() {
		const data = getData();
		originalData = { ...data }; // 保存原始数据
		
		// 检查数据是否为空
		if (checkDataEmpty(data)) {
			return;
		}
		
		// 设置默认日期范围（最近30天）
		const endDate = new Date();
		const startDate = new Date();
		startDate.setDate(startDate.getDate() - 30);
		
		const startDateInput = document.getElementById("startDate");
		const endDateInput = document.getElementById("endDate");
		if (startDateInput) {
			startDateInput.value = startDate.toISOString().split('T')[0];
		}
		if (endDateInput) {
			endDateInput.value = endDate.toISOString().split('T')[0];
		}
		
		// 应用初始筛选
		const filteredData = filterByDateRange(data, startDateInput?.value, endDateInput?.value);
		
		renderSummary(filteredData);
		const agg = aggregateBySales(filteredData);
		currentSalesData = agg;
		renderTable("salesAgg", agg);
		
		// 绑定时间范围筛选
		const applyFilterBtn = document.getElementById("applyFilterBtn");
		const resetFilterBtn = document.getElementById("resetFilterBtn");
		
		if (applyFilterBtn) {
			applyFilterBtn.addEventListener("click", () => {
				const start = startDateInput?.value;
				const end = endDateInput?.value;
				const filtered = filterByDateRange(originalData, start, end);
				renderSummary(filtered);
				const newAgg = aggregateBySales(filtered);
				currentSalesData = newAgg;
				renderTable("salesAgg", newAgg);
				drawBarChart("revenueRefundChart", newAgg);
				drawPieChart("netRevenueChart", newAgg);
				
				// 重新渲染所有分析和报告
				renderAllAnalysis(filtered);
				const localAnalysis = generateLocalAnalysis(filtered);
				document.getElementById("overviewAnalysis").innerHTML = localAnalysis.overview;
				document.getElementById("financialAnalysis").innerHTML = localAnalysis.financial;
				document.getElementById("marketingAnalysis").innerHTML = localAnalysis.marketing;
			});
		}
		
		if (resetFilterBtn) {
			resetFilterBtn.addEventListener("click", () => {
				if (startDateInput) startDateInput.value = "";
				if (endDateInput) endDateInput.value = "";
				renderSummary(originalData);
				const newAgg = aggregateBySales(originalData);
				currentSalesData = newAgg;
				renderTable("salesAgg", newAgg);
				drawBarChart("revenueRefundChart", newAgg);
				drawPieChart("netRevenueChart", newAgg);
				
				// 重新渲染所有分析和报告
				renderAllAnalysis(originalData);
				const localAnalysis = generateLocalAnalysis(originalData);
				document.getElementById("overviewAnalysis").innerHTML = localAnalysis.overview;
				document.getElementById("financialAnalysis").innerHTML = localAnalysis.financial;
				document.getElementById("marketingAnalysis").innerHTML = localAnalysis.marketing;
			});
		}
		
		// 绑定导出按钮
		const exportBtn = document.getElementById("exportDataBtn");
		if (exportBtn) {
			exportBtn.addEventListener("click", () => {
				exportAnalysisResults(currentSalesData);
			});
		}
		
		// 绑定多选模式切换
		const multiSelectCheckbox = document.getElementById("multiSelectMode");
		if (multiSelectCheckbox) {
			multiSelectCheckbox.addEventListener("change", (e) => {
				isMultiSelectMode = e.target.checked;
				if (!isMultiSelectMode) {
					// 关闭多选模式时，如果有待应用的选择，自动应用
					if (pendingSelections.size > 0) {
						applyPendingSelections(agg);
					}
					pendingSelections.clear();
					updatePendingInfo();
				}
			});
		}
		
		// 绑定应用筛选按钮
		const applyBtn = document.getElementById("applyFilter");
		if (applyBtn) {
			applyBtn.addEventListener("click", () => {
				applyPendingSelections(agg);
			});
		}
		
		// 绘制图表
		setTimeout(() => {
			drawBarChart("revenueRefundChart", agg);
			drawPieChart("netRevenueChart", agg);
			
			// 渲染所有多维度分析
			renderAllAnalysis(filteredData);
			
			// 生成本地智能分析报告
			const localAnalysis = generateLocalAnalysis(filteredData);
			document.getElementById("overviewAnalysis").innerHTML = localAnalysis.overview;
			document.getElementById("financialAnalysis").innerHTML = localAnalysis.financial;
			document.getElementById("marketingAnalysis").innerHTML = localAnalysis.marketing;
		}, 100);
	}

	// 工具函数：颜色变亮
	function lightenColor(color, percent) {
		const num = parseInt(color.replace("#", ""), 16);
		const r = Math.min(255, (num >> 16) + percent);
		const g = Math.min(255, ((num >> 8) & 0x00FF) + percent);
		const b = Math.min(255, (num & 0x0000FF) + percent);
		return `#${((r << 16) | (g << 8) | b).toString(16).padStart(6, "0")}`;
	}

	// 选项卡切换
	function initTabs() {
		const tabBtns = document.querySelectorAll(".tab-btn");
		const tabContents = document.querySelectorAll(".tab-content");
		
		tabBtns.forEach(btn => {
			btn.addEventListener("click", () => {
				const targetTab = btn.getAttribute("data-tab");
				
				// 移除所有活动状态
				tabBtns.forEach(b => b.classList.remove("active"));
				tabContents.forEach(c => {
					c.classList.remove("active");
					c.style.display = "none";
				});
				
				// 激活当前选项卡
				btn.classList.add("active");
				const targetContent = document.getElementById(targetTab + "Tab");
				if (targetContent) {
					targetContent.classList.add("active");
					targetContent.style.display = "block";
				}
			});
		});
	}

	// 按季度分析营收
	function analyzeByQuarter(data) {
		const quarterMap = new Map();
		
		data.revenue.forEach(row => {
			const quarter = row["季度"] || "未知季度";
			const amount = Number(row["课程拆分金额"] || 0);
			
			if (!quarterMap.has(quarter)) {
				quarterMap.set(quarter, { revenue: 0, refund: 0, count: 0 });
			}
			const q = quarterMap.get(quarter);
			q.revenue += amount;
			q.count += 1;
		});
		
		// 计算退费
		data.refund.forEach(row => {
			const quarter = row["季度"] || row["学期"] || "未知季度";
			const amount = Number(row["退费金额"] || 0);
			
			if (!quarterMap.has(quarter)) {
				quarterMap.set(quarter, { revenue: 0, refund: 0, count: 0 });
			}
			const q = quarterMap.get(quarter);
			q.refund += amount;
		});
		
		return Array.from(quarterMap.entries()).map(([quarter, data]) => ({
			季度: quarter,
			营收: Number(data.revenue.toFixed(2)),
			退费: Number(data.refund.toFixed(2)),
			净营收: Number((data.revenue - data.refund).toFixed(2)),
			订单数: data.count
		})).sort((a, b) => a.季度.localeCompare(b.季度));
	}

	// 按学科分析营收
	function analyzeBySubject(data) {
		const subjectMap = new Map();
		
		data.revenue.forEach(row => {
			const subject = row["学科"] || "未知学科";
			const amount = Number(row["课程拆分金额"] || 0);
			
			if (!subjectMap.has(subject)) {
				subjectMap.set(subject, { revenue: 0, refund: 0, count: 0 });
			}
			const s = subjectMap.get(subject);
			s.revenue += amount;
			s.count += 1;
		});
		
		data.refund.forEach(row => {
			const subject = row["科目"] || "未知学科";
			const amount = Number(row["退费金额"] || 0);
			
			if (!subjectMap.has(subject)) {
				subjectMap.set(subject, { revenue: 0, refund: 0, count: 0 });
			}
			const s = subjectMap.get(subject);
			s.refund += amount;
		});
		
		return Array.from(subjectMap.entries()).map(([subject, data]) => ({
			学科: subject,
			营收: Number(data.revenue.toFixed(2)),
			退费: Number(data.refund.toFixed(2)),
			净营收: Number((data.revenue - data.refund).toFixed(2)),
			订单数: data.count
		})).sort((a, b) => b.营收 - a.营收);
	}

	// 各学科不同周期营收分析
	function analyzeSubjectByPeriod(data) {
		const matrix = new Map(); // 学科 -> 周期 -> 数据
		
		data.revenue.forEach(row => {
			const subject = row["学科"] || "未知学科";
			const period = row["季度"] || row["班期"] || "未知周期";
			const amount = Number(row["课程拆分金额"] || 0);
			
			if (!matrix.has(subject)) {
				matrix.set(subject, new Map());
			}
			const subjectData = matrix.get(subject);
			
			if (!subjectData.has(period)) {
				subjectData.set(period, { revenue: 0, count: 0 });
			}
			const periodData = subjectData.get(period);
			periodData.revenue += amount;
			periodData.count += 1;
		});
		
		// 转换为表格数据
		const result = [];
		const allPeriods = new Set();
		
		matrix.forEach((periodData, subject) => {
			periodData.forEach((data, period) => {
				allPeriods.add(period);
			});
		});
		
		matrix.forEach((periodData, subject) => {
			const row = { 学科: subject };
			allPeriods.forEach(period => {
				const data = periodData.get(period) || { revenue: 0, count: 0 };
				row[period] = Number(data.revenue.toFixed(2));
			});
			result.push(row);
		});
		
		return { data: result, periods: Array.from(allPeriods).sort() };
	}

	// 按班型分析
	function analyzeByClassType(data) {
		const classTypeMap = new Map();
		
		data.revenue.forEach(row => {
			const classType = row["班型"] || "未知班型";
			const amount = Number(row["课程拆分金额"] || 0);
			
			if (!classTypeMap.has(classType)) {
				classTypeMap.set(classType, { revenue: 0, count: 0 });
			}
			const ct = classTypeMap.get(classType);
			ct.revenue += amount;
			ct.count += 1;
		});
		
		return Array.from(classTypeMap.entries()).map(([classType, data]) => ({
			班型: classType,
			营收: Number(data.revenue.toFixed(2)),
			订单数: data.count,
			平均订单金额: Number((data.revenue / data.count).toFixed(2))
		})).sort((a, b) => b.营收 - a.营收);
	}

	// 按学生类型分析
	function analyzeByStudentType(data) {
		const studentTypeMap = new Map();
		
		data.revenue.forEach(row => {
			const studentType = row["学生类型"] || "未知类型";
			const amount = Number(row["课程拆分金额"] || 0);
			
			if (!studentTypeMap.has(studentType)) {
				studentTypeMap.set(studentType, { revenue: 0, refund: 0, count: 0 });
			}
			const st = studentTypeMap.get(studentType);
			st.revenue += amount;
			st.count += 1;
		});
		
		data.refund.forEach(row => {
			const studentType = row["是否新生"] === "是" ? "新生" : (row["是否新生"] === "否" ? "老生" : "未知类型");
			const amount = Number(row["退费金额"] || 0);
			
			if (!studentTypeMap.has(studentType)) {
				studentTypeMap.set(studentType, { revenue: 0, refund: 0, count: 0 });
			}
			const st = studentTypeMap.get(studentType);
			st.refund += amount;
		});
		
		return Array.from(studentTypeMap.entries()).map(([studentType, data]) => ({
			学生类型: studentType,
			营收: Number(data.revenue.toFixed(2)),
			退费: Number(data.refund.toFixed(2)),
			净营收: Number((data.revenue - data.refund).toFixed(2)),
			订单数: data.count,
			退费率: data.revenue > 0 ? Number((data.refund / data.revenue * 100).toFixed(2)) : 0
		})).sort((a, b) => b.营收 - a.营收);
	}

	// 营收趋势分析（按时间）
	function analyzeRevenueTrend(data) {
		const dateMap = new Map();
		
		data.revenue.forEach(row => {
			const dateStr = row["收款日期"] || row["订单创建时间"] || "";
			if (!dateStr) return;
			
			const date = new Date(dateStr);
			if (isNaN(date.getTime())) return;
			
			const dateKey = date.toISOString().split('T')[0];
			const amount = Number(row["课程拆分金额"] || 0);
			
			if (!dateMap.has(dateKey)) {
				dateMap.set(dateKey, { revenue: 0, refund: 0, count: 0 });
			}
			const d = dateMap.get(dateKey);
			d.revenue += amount;
			d.count += 1;
		});
		
		data.refund.forEach(row => {
			const dateStr = row["收款日期"] || row["申请退款日期"] || "";
			if (!dateStr) return;
			
			const date = new Date(dateStr);
			if (isNaN(date.getTime())) return;
			
			const dateKey = date.toISOString().split('T')[0];
			const amount = Number(row["退费金额"] || 0);
			
			if (!dateMap.has(dateKey)) {
				dateMap.set(dateKey, { revenue: 0, refund: 0, count: 0 });
			}
			const d = dateMap.get(dateKey);
			d.refund += amount;
		});
		
		return Array.from(dateMap.entries())
			.map(([date, data]) => ({
				日期: date,
				营收: Number(data.revenue.toFixed(2)),
				退费: Number(data.refund.toFixed(2)),
				净营收: Number((data.revenue - data.refund).toFixed(2)),
				订单数: data.count
			}))
			.sort((a, b) => a.日期.localeCompare(b.日期));
	}

	// 绘制营收趋势图
	function drawRevenueTrendChart(canvasId, trendData) {
		const canvas = document.getElementById(canvasId);
		if (!canvas || !trendData || trendData.length === 0) return;
		
		const ctx = canvas.getContext("2d");
		const width = canvas.width;
		const height = canvas.height;
		const padding = { top: 40, right: 40, bottom: 60, left: 80 };
		const chartWidth = width - padding.left - padding.right;
		const chartHeight = height - padding.top - padding.bottom;
		
		ctx.clearRect(0, 0, width, height);
		
		const maxValue = Math.max(...trendData.map(d => Math.max(d.营收, d.退费)));
		const maxY = Math.ceil(maxValue * 1.1);
		
		// 绘制坐标轴
		ctx.strokeStyle = "#4b5563";
		ctx.lineWidth = 1;
		ctx.beginPath();
		ctx.moveTo(padding.left, padding.top);
		ctx.lineTo(padding.left, padding.top + chartHeight);
		ctx.lineTo(padding.left + chartWidth, padding.top + chartHeight);
		ctx.stroke();
		
		// Y轴刻度
		ctx.fillStyle = "#9ca3af";
		ctx.font = "11px sans-serif";
		ctx.textAlign = "right";
		for (let i = 0; i <= 5; i++) {
			const value = (maxY / 5) * i;
			const y = padding.top + chartHeight - (value / maxY) * chartHeight;
			ctx.fillText(value.toFixed(0), padding.left - 10, y + 4);
		}
		
		// 绘制折线（旧版，保留兼容性）
		const pointSpacing = trendData.length > 1 ? chartWidth / (trendData.length - 1) : chartWidth;
		
		// 营收折线
		ctx.beginPath();
		ctx.strokeStyle = "#3b82f6";
		ctx.lineWidth = 2;
		trendData.forEach((d, idx) => {
			const x = padding.left + idx * pointSpacing;
			const y = padding.top + chartHeight - (d.营收 / maxY) * chartHeight;
			if (idx === 0) {
				ctx.moveTo(x, y);
			} else {
				ctx.lineTo(x, y);
			}
			// 绘制点
			ctx.fillStyle = "#3b82f6";
			ctx.beginPath();
			ctx.arc(x, y, 3, 0, Math.PI * 2);
			ctx.fill();
		});
		ctx.stroke();
		
		// 退费折线
		ctx.beginPath();
		ctx.strokeStyle = "#ef4444";
		ctx.lineWidth = 2;
		trendData.forEach((d, idx) => {
			const x = padding.left + idx * pointSpacing;
			const y = padding.top + chartHeight - (d.退费 / maxY) * chartHeight;
			if (idx === 0) {
				ctx.moveTo(x, y);
			} else {
				ctx.lineTo(x, y);
			}
			// 绘制点
			ctx.fillStyle = "#ef4444";
			ctx.beginPath();
			ctx.arc(x, y, 3, 0, Math.PI * 2);
			ctx.fill();
		});
		ctx.stroke();
		
		// 净营收折线（用于净营收趋势图）
		if (canvasId === "netRevenueTrendChart") {
			ctx.beginPath();
			ctx.strokeStyle = "#10b981";
			ctx.lineWidth = 2;
			trendData.forEach((d, idx) => {
				const x = padding.left + idx * pointSpacing;
				const y = padding.top + chartHeight - (d.净营收 / maxY) * chartHeight;
				if (idx === 0) {
					ctx.moveTo(x, y);
				} else {
					ctx.lineTo(x, y);
				}
				// 绘制点
				ctx.fillStyle = "#10b981";
				ctx.beginPath();
				ctx.arc(x, y, 3, 0, Math.PI * 2);
				ctx.fill();
			});
			ctx.stroke();
		}
		
		// 图例
		ctx.fillStyle = "#3b82f6";
		ctx.fillRect(padding.left + chartWidth - 100, padding.top - 25, 12, 12);
		ctx.fillStyle = "#e5e7eb";
		ctx.font = "11px sans-serif";
		ctx.textAlign = "left";
		ctx.fillText("营收", padding.left + chartWidth - 85, padding.top - 15);
		
		ctx.fillStyle = "#ef4444";
		ctx.fillRect(padding.left + chartWidth - 50, padding.top - 25, 12, 12);
		ctx.fillText("退费", padding.left + chartWidth - 35, padding.top - 15);
		
		if (canvasId === "netRevenueTrendChart") {
			ctx.fillStyle = "#10b981";
			ctx.fillRect(padding.left + chartWidth - 150, padding.top - 25, 12, 12);
			ctx.fillText("净营收", padding.left + chartWidth - 135, padding.top - 15);
		}
	}

	// 绘制退费率图表
	function drawRefundRateChart(canvasId, data) {
		const canvas = document.getElementById(canvasId);
		if (!canvas || !data || data.length === 0) return;
		
		const ctx = canvas.getContext("2d");
		// 清除画布
		ctx.clearRect(0, 0, canvas.width, canvas.height);
		const width = canvas.width;
		const height = canvas.height;
		const padding = { top: 40, right: 40, bottom: 60, left: 80 };
		const chartWidth = width - padding.left - padding.right;
		const chartHeight = height - padding.top - padding.bottom;
		
		ctx.clearRect(0, 0, width, height);
		
		// 计算退费率
		const refundRates = data.map(d => ({
			label: d.季度 || d.学科 || "未知",
			退费率: d.营收 > 0 ? (d.退费 / d.营收 * 100) : 0
		}));
		
		const maxRate = Math.max(...refundRates.map(r => r.退费率), 10);
		const barWidth = chartWidth / refundRates.length * 0.8;
		const gap = chartWidth / refundRates.length * 0.2;
		
		// 绘制坐标轴
		ctx.strokeStyle = "#4b5563";
		ctx.lineWidth = 1;
		ctx.beginPath();
		ctx.moveTo(padding.left, padding.top);
		ctx.lineTo(padding.left, padding.top + chartHeight);
		ctx.lineTo(padding.left + chartWidth, padding.top + chartHeight);
		ctx.stroke();
		
		// Y轴刻度
		ctx.fillStyle = "#9ca3af";
		ctx.font = "11px sans-serif";
		ctx.textAlign = "right";
		for (let i = 0; i <= 5; i++) {
			const value = (maxRate / 5) * i;
			const y = padding.top + chartHeight - (value / maxRate) * chartHeight;
			ctx.fillText(value.toFixed(1) + "%", padding.left - 10, y + 4);
		}
		
		// 绘制柱状图
		refundRates.forEach((rate, idx) => {
			const x = padding.left + idx * (barWidth + gap) + gap / 2;
			const barHeight = (rate.退费率 / maxRate) * chartHeight;
			const color = rate.退费率 > 10 ? "#ef4444" : rate.退费率 > 5 ? "#f59e0b" : "#10b981";
			
			ctx.fillStyle = color;
			ctx.fillRect(x, padding.top + chartHeight - barHeight, barWidth, barHeight);
			
			// 标签
			ctx.fillStyle = "#e5e7eb";
			ctx.font = "10px sans-serif";
			ctx.textAlign = "center";
			ctx.save();
			ctx.translate(x + barWidth / 2, padding.top + chartHeight + 15);
			ctx.rotate(-Math.PI / 4);
			ctx.fillText(rate.label.length > 6 ? rate.label.substring(0, 6) + "..." : rate.label, 0, 0);
			ctx.restore();
			
			// 数值
			ctx.fillStyle = "#e5e7eb";
			ctx.font = "10px sans-serif";
			ctx.textAlign = "center";
			ctx.fillText(rate.退费率.toFixed(1) + "%", x + barWidth / 2, padding.top + chartHeight - barHeight - 5);
		});
	}

	// 绘制学科×周期柱状图
	function drawSubjectPeriodChart(canvasId, subjectPeriodData) {
		const canvas = document.getElementById(canvasId);
		if (!canvas || !subjectPeriodData || subjectPeriodData.data.length === 0) return;
		
		const ctx = canvas.getContext("2d");
		// 清除画布
		ctx.clearRect(0, 0, canvas.width, canvas.height);
		const width = canvas.width;
		const height = canvas.height;
		// 为竖版图例预留右侧空间
		const legendWidth = 120;
		const padding = { top: 40, right: legendWidth + 20, bottom: 80, left: 80 };
		const chartWidth = width - padding.left - padding.right;
		const chartHeight = height - padding.top - padding.bottom;
		
		ctx.clearRect(0, 0, width, height);
		
		const { data, periods } = subjectPeriodData;
		const maxValue = Math.max(...data.flatMap(row => 
			periods.map(p => Number(row[p] || 0))
		));
		const maxY = Math.ceil(maxValue * 1.1);
		
		// 绘制坐标轴
		ctx.strokeStyle = "#4b5563";
		ctx.lineWidth = 1;
		ctx.beginPath();
		ctx.moveTo(padding.left, padding.top);
		ctx.lineTo(padding.left, padding.top + chartHeight);
		ctx.lineTo(padding.left + chartWidth, padding.top + chartHeight);
		ctx.stroke();
		
		// Y轴刻度
		ctx.fillStyle = "#9ca3af";
		ctx.font = "11px sans-serif";
		ctx.textAlign = "right";
		for (let i = 0; i <= 5; i++) {
			const value = (maxY / 5) * i;
			const y = padding.top + chartHeight - (value / maxY) * chartHeight;
			ctx.fillText(value.toFixed(0), padding.left - 10, y + 4);
		}
		
		// 颜色方案
		const colors = [
			"#3b82f6", "#10b981", "#f59e0b", "#ef4444", "#8b5cf6",
			"#06b6d4", "#ec4899", "#84cc16", "#f97316", "#6366f1"
		];
		
		// 绘制分组柱状图
		const groupWidth = chartWidth / data.length;
		const barWidth = groupWidth / (periods.length + 1) * 0.8;
		const gap = groupWidth / (periods.length + 1) * 0.2;
		
		data.forEach((row, rowIdx) => {
			const groupX = padding.left + rowIdx * groupWidth;
			
			periods.forEach((period, periodIdx) => {
				const value = Number(row[period] || 0);
				const barHeight = (value / maxY) * chartHeight;
				const x = groupX + periodIdx * (barWidth + gap) + gap;
				const color = colors[periodIdx % colors.length];
				
				ctx.fillStyle = color;
				ctx.fillRect(x, padding.top + chartHeight - barHeight, barWidth, barHeight);
			});
			
			// X轴标签（学科名）
			ctx.fillStyle = "#e5e7eb";
			ctx.font = "10px sans-serif";
			ctx.textAlign = "center";
			ctx.save();
			ctx.translate(groupX + groupWidth / 2, padding.top + chartHeight + 20);
			ctx.rotate(-Math.PI / 4);
			const label = row.学科.length > 6 ? row.学科.substring(0, 6) + "..." : row.学科;
			ctx.fillText(label, 0, 0);
			ctx.restore();
		});
		
		// 图例（竖版排列，避免显示不齐）
		let legendX = padding.left + chartWidth + 10;
		let legendY = padding.top;
		const legendItemHeight = 18;
		const maxLegendWidth = 100;
		
		periods.forEach((period, idx) => {
			ctx.fillStyle = colors[idx % colors.length];
			ctx.fillRect(legendX, legendY, 12, 12);
			ctx.fillStyle = "#e5e7eb";
			ctx.font = "10px sans-serif";
			ctx.textAlign = "left";
			const label = period.length > 10 ? period.substring(0, 10) + "..." : period;
			ctx.fillText(label, legendX + 15, legendY + 9);
			legendY += legendItemHeight;
		});
	}

	// 渲染表格
	function renderTableData(containerId, data, headers = null) {
		const container = document.getElementById(containerId);
		if (!container) return;
		
		container.innerHTML = "";
		if (!data || data.length === 0) {
			container.textContent = "暂无数据";
			return;
		}
		
		const table = document.createElement("table");
		const thead = document.createElement("thead");
		const trh = document.createElement("tr");
		
		const tableHeaders = headers || Object.keys(data[0]);
		tableHeaders.forEach(h => {
			const th = document.createElement("th");
			th.textContent = h;
			trh.appendChild(th);
		});
		thead.appendChild(trh);
		table.appendChild(thead);
		
		const tbody = document.createElement("tbody");
		data.forEach(row => {
			const tr = document.createElement("tr");
			tableHeaders.forEach(h => {
				const td = document.createElement("td");
				const val = row[h];
				if (typeof val === 'number') {
					td.textContent = val.toFixed(2);
					td.style.textAlign = "right";
				} else {
					td.textContent = val || "";
				}
				tr.appendChild(td);
			});
			tbody.appendChild(tr);
		});
		
		table.appendChild(tbody);
		container.appendChild(table);
	}

	// 渲染所有分析
	function renderAllAnalysis(data) {
		// 财务分析
		const quarterData = analyzeByQuarter(data);
		renderTableData("quarterRevenueTable", quarterData);
		
		const trendData = analyzeRevenueTrend(data);
		if (trendData.length > 0) {
			// 使用增强版趋势图
			const showMovingAvg = document.getElementById("showMovingAverage")?.checked ?? true;
			drawRevenueTrendChartEnhanced("revenueTrendChart", trendData, showMovingAvg);
			drawRevenueTrendChartEnhanced("netRevenueTrendChart", trendData, showMovingAvg);
			drawRefundRateChart("refundRateChart", quarterData);
			
			// 生成趋势图洞察
			const trendInsights = generateChartInsights("revenueTrend", trendData);
			const insightsEl = document.getElementById("revenueTrendInsights");
			if (insightsEl && trendInsights.length > 0) {
				insightsEl.innerHTML = trendInsights.map(i => `<div class="insight-item">${i}</div>`).join('');
			}
			
			// 生成退费率洞察
			const refundInsights = [];
			quarterData.forEach(q => {
				const rate = q.营收 > 0 ? (q.退费 / q.营收 * 100) : 0;
				if (rate > 10) {
					refundInsights.push(`${q.季度}退费率${rate.toFixed(1)}%，需要关注`);
				}
			});
			const refundInsightsEl = document.getElementById("refundRateInsights");
			if (refundInsightsEl && refundInsights.length > 0) {
				refundInsightsEl.innerHTML = refundInsights.map(i => `<div class="insight-item">${i}</div>`).join('');
			}
			
			// 季度洞察
			const quarterInsights = [];
			if (quarterData.length >= 2) {
				const best = quarterData.reduce((max, q) => q.营收 > max.营收 ? q : max, quarterData[0]);
				const worst = quarterData.reduce((min, q) => q.营收 < min.营收 ? q : min, quarterData[0]);
				quarterInsights.push(`最佳季度：${best.季度}（${formatToWanYuan(best.营收)}）`);
				quarterInsights.push(`最弱季度：${worst.季度}（${formatToWanYuan(worst.营收)}）`);
			}
			const quarterInsightsEl = document.getElementById("quarterInsights");
			if (quarterInsightsEl && quarterInsights.length > 0) {
				quarterInsightsEl.innerHTML = quarterInsights.map(i => `<div class="insight-item">${i}</div>`).join('');
			}
		}
		
		// 营销分析
		const subjectData = analyzeBySubject(data);
		drawPieChart("subjectRevenueChart", subjectData.map(s => ({
			销售人员: s.学科,
			营收: s.营收,
			退费金额: s.退费,
			净营收: s.净营收
		})));
		
		const subjectPeriodData = analyzeSubjectByPeriod(data);
		renderTableData("subjectPeriodMatrix", subjectPeriodData.data, ["学科", ...subjectPeriodData.periods]);
		
		// 绘制学科×周期柱状图（使用更大的画布以容纳竖版图例）
		if (subjectPeriodData.data.length > 0 && subjectPeriodData.periods.length > 0) {
			const canvas = document.getElementById("subjectPeriodChart");
			if (canvas) {
				// 根据周期数量调整画布宽度，为图例留出空间
				const legendWidth = 120;
				canvas.width = 600 + legendWidth;
			}
			drawSubjectPeriodChart("subjectPeriodChart", subjectPeriodData);
		}
		
		const classTypeData = analyzeByClassType(data);
		drawPieChart("classTypeChart", classTypeData.map(ct => ({
			销售人员: ct.班型,
			营收: ct.营收,
			退费金额: 0,
			净营收: ct.营收
		})));
		
		const studentTypeData = analyzeByStudentType(data);
		drawPieChart("studentTypeChart", studentTypeData.map(st => ({
			销售人员: st.学生类型,
			营收: st.营收,
			退费金额: st.退费,
			净营收: st.净营收
		})));
	}

	// ==================== 智能文字分析功能 ====================
	
	// 数据脱敏（移除敏感信息，只保留统计摘要）
	function sanitizeDataForAI(data) {
		return {
			summary: {
				totalRevenue: data.revenue.reduce((sum, r) => sum + (Number(r["课程拆分金额"]) || 0), 0),
				totalRefund: data.refund.reduce((sum, r) => sum + (Number(r["退费金额"]) || 0), 0),
				revenueCount: data.revenue.length,
				refundCount: data.refund.length
			},
			byQuarter: analyzeByQuarter(data).map(q => ({
				季度: q.季度,
				营收: q.营收,
				退费: q.退费,
				净营收: q.净营收
			})),
			bySubject: analyzeBySubject(data).map(s => ({
				学科: s.学科,
				营收: s.营收,
				退费: s.退费,
				净营收: s.净营收
			})),
			byClassType: analyzeByClassType(data).map(ct => ({
				班型: ct.班型,
				营收: ct.营收,
				订单数: ct.订单数
			})),
			byStudentType: analyzeByStudentType(data).map(st => ({
				学生类型: st.学生类型,
				营收: st.营收,
				退费率: st.退费率
			}))
		};
	}

	// 本地智能分析（不依赖外部API，数据安全）
	function generateLocalAnalysis(data) {
		const analysis = {
			overview: generateOverviewAnalysis(data),
			financial: generateFinancialAnalysis(data),
			marketing: generateMarketingAnalysis(data)
		};
		return analysis;
	}

	// 生成概览分析（优化排版）
	function generateOverviewAnalysis(data) {
		const totalRevenue = data.revenue.reduce((sum, r) => sum + (Number(r["课程拆分金额"]) || 0), 0);
		const totalRefund = data.refund.reduce((sum, r) => sum + (Number(r["退费金额"]) || 0), 0);
		const netRevenue = totalRevenue - totalRefund;
		const refundRate = totalRevenue > 0 ? (totalRefund / totalRevenue * 100) : 0;
		const revenueCount = data.revenue.length;
		const refundCount = data.refund.length;
		const avgOrderValue = revenueCount > 0 ? totalRevenue / revenueCount : 0;
		const netRevenueRate = totalRevenue > 0 ? (netRevenue / totalRevenue * 100) : 0;
		
		// 关键指标卡片
		let html = `<div class="key-metrics">`;
		html += `<div class="metric-card ${refundRate < 5 ? 'success' : refundRate < 10 ? 'warning' : 'danger'}">`;
		html += `<div class="metric-label">退费率</div>`;
		html += `<div class="metric-value">${refundRate.toFixed(1)}%</div>`;
		html += `<div class="metric-trend">${refundRate < 5 ? '健康' : refundRate < 10 ? '需关注' : '偏高'}</div>`;
		html += `</div>`;
		
		html += `<div class="metric-card success">`;
		html += `<div class="metric-label">净营收率</div>`;
		html += `<div class="metric-value">${netRevenueRate.toFixed(1)}%</div>`;
		html += `<div class="metric-trend">${netRevenueRate > 70 ? '优秀' : netRevenueRate > 50 ? '良好' : '偏低'}</div>`;
		html += `</div>`;
		
		html += `<div class="metric-card">`;
		html += `<div class="metric-label">平均订单</div>`;
		html += `<div class="metric-value">${avgOrderValue.toFixed(0)}元</div>`;
		html += `<div class="metric-trend">${revenueCount}笔订单</div>`;
		html += `</div>`;
		
		html += `<div class="metric-card">`;
		html += `<div class="metric-label">净营收</div>`;
		html += `<div class="metric-value">${formatToWanYuan(netRevenue)}</div>`;
		html += `<div class="metric-trend">${formatToWanYuan(totalRevenue)} - ${formatToWanYuan(totalRefund)}</div>`;
		html += `</div>`;
		html += `</div>`;
		
		// 重点提示
		if (refundRate >= 10) {
			html += `<div class="insight-box danger">`;
			html += `<div class="insight-box-title">⚠️ 退费率偏高，需要立即关注</div>`;
			html += `<div class="insight-box-content">当前退费率 ${refundRate.toFixed(1)}% 超过健康水平，建议：分析退费原因、加强课程质量管控、提升客户服务体验。</div>`;
			html += `</div>`;
		} else if (refundRate >= 5) {
			html += `<div class="insight-box warning">`;
			html += `<div class="insight-box-title">⚠️ 退费率需关注</div>`;
			html += `<div class="insight-box-content">当前退费率 ${refundRate.toFixed(1)}% 处于可接受范围，建议持续关注退费原因，优化课程和服务质量。</div>`;
			html += `</div>`;
		} else {
			html += `<div class="insight-box success">`;
			html += `<div class="insight-box-title">✓ 退费率健康</div>`;
			html += `<div class="insight-box-content">当前退费率 ${refundRate.toFixed(1)}% 处于健康水平，说明客户满意度较高，课程质量稳定。</div>`;
			html += `</div>`;
		}
		
		// 业务概况（简洁版）
		html += `<div class="analysis-section">`;
		html += `<div class="analysis-section-title">📊 业务概况</div>`;
		html += `<div class="analysis-text">`;
		html += `<p>总营收 <span class="highlight">${formatToWanYuan(totalRevenue)}</span>，净营收 <span class="highlight">${formatToWanYuan(netRevenue)}</span>，共 <span class="highlight">${revenueCount}</span> 笔订单。</p>`;
		html += `</div></div>`;
		
		// 添加概览建议
		html += `<div class="insight-box info" style="margin-top: 16px;">`;
		html += `<div class="insight-box-title">💡 核心建议</div>`;
		html += `<ul class="suggestions-list">`;
		
		if (refundRate >= 10) {
			html += `<li class="urgent">退费率${refundRate.toFixed(1)}%严重偏高，需立即建立退费专项小组，深入分析退费原因</li>`;
			html += `<li class="urgent">建立退费预警机制，设定退费率"红线"指标（建议<10%）</li>`;
			html += `<li>加强课程质量管控，提升客户服务体验</li>`;
		} else if (refundRate >= 5) {
			html += `<li class="important">退费率${refundRate.toFixed(1)}%需持续关注，建议建立退费原因追踪体系</li>`;
			html += `<li>优化课程内容匹配度，提升客户满意度</li>`;
		} else {
			html += `<li>退费率${refundRate.toFixed(1)}%处于健康水平，继续保持课程质量和服务标准</li>`;
		}
		
		if (netRevenueRate < 50) {
			html += `<li class="important">净营收率${netRevenueRate.toFixed(1)}%偏低（目标>70%），需优化成本结构，提高盈利能力</li>`;
		}
		
		if (revenueCount > 0) {
			const avgOrderValue = totalRevenue / revenueCount;
			if (avgOrderValue < 2000) {
				html += `<li>平均订单金额${avgOrderValue.toFixed(0)}元偏低，建议优化课程定价策略，提升客单价</li>`;
			}
		}
		
		html += `<li>建立完整的数据追踪体系，定期进行业务健康度评估</li>`;
		html += `<li>关注客户生命周期价值，提升客户留存和复购率</li>`;
		html += `</ul></div>`;
		
		return html;
	}

	// 生成财务分析（优化排版）
	function generateFinancialAnalysis(data) {
		const quarterData = analyzeByQuarter(data);
		const trendData = analyzeRevenueTrend(data);
		
		let html = `<div class="analysis-content">`;
		
		// 季度分析卡片
		if (quarterData.length > 0) {
			const bestQuarter = quarterData.reduce((max, q) => q.营收 > max.营收 ? q : max, quarterData[0]);
			const worstQuarter = quarterData.reduce((min, q) => q.营收 < min.营收 ? q : min, quarterData[0]);
			
			html += `<div class="analysis-section">`;
			html += `<div class="analysis-section-title">📅 季度表现</div>`;
			html += `<div class="analysis-text">`;
			html += `<p>最佳：<span class="success">${bestQuarter.季度}</span> ${formatToWanYuan(bestQuarter.营收)}</p>`;
			html += `<p>最弱：<span class="warning">${worstQuarter.季度}</span> ${formatToWanYuan(worstQuarter.营收)}</p>`;
			
			// 季度增长分析
			if (quarterData.length >= 2) {
				const sorted = [...quarterData].sort((a, b) => a.季度.localeCompare(b.季度));
				for (let i = 1; i < sorted.length; i++) {
					const growth = sorted[i].营收 - sorted[i-1].营收;
					const growthRate = sorted[i-1].营收 > 0 ? (growth / sorted[i-1].营收 * 100) : 0;
					if (Math.abs(growthRate) > 5) {
						html += `<p>${sorted[i].季度} vs ${sorted[i-1].季度}：<span class="${growthRate > 0 ? 'success' : 'warning'}">${growthRate > 0 ? '+' : ''}${growthRate.toFixed(1)}%</span></p>`;
					}
				}
			}
			html += `</div></div>`;
		}
		
		// 趋势分析卡片
		if (trendData.length >= 7) {
			const recent = trendData.slice(-7);
			const earlier = trendData.slice(0, Math.min(7, trendData.length - 7));
			const recentAvg = recent.reduce((sum, d) => sum + d.营收, 0) / recent.length;
			const earlierAvg = earlier.length > 0 ? earlier.reduce((sum, d) => sum + d.营收, 0) / earlier.length : recentAvg;
			const trend = recentAvg - earlierAvg;
			const trendRate = earlierAvg > 0 ? (trend / earlierAvg * 100) : 0;
			
			html += `<div class="analysis-section">`;
			html += `<div class="analysis-section-title">📈 近期趋势</div>`;
			html += `<div class="analysis-text">`;
			
			if (trendRate > 10) {
				html += `<p><span class="success">强劲上升</span> ${trendRate.toFixed(1)}%</p>`;
			} else if (trendRate > 0) {
				html += `<p><span class="success">温和上升</span> ${trendRate.toFixed(1)}%</p>`;
			} else if (trendRate > -10) {
				html += `<p><span class="warning">小幅下降</span> ${Math.abs(trendRate).toFixed(1)}%</p>`;
			} else {
				html += `<p><span class="danger">明显下降</span> ${Math.abs(trendRate).toFixed(1)}%</p>`;
			}
			html += `<p>最近7天 vs 前期</p>`;
			html += `</div></div>`;
			
			// 如果下降明显，显示建议
			if (trendRate <= -10) {
				html += `<div class="insight-box danger">`;
				html += `<div class="insight-box-title">⚠️ 需要立即采取行动</div>`;
				html += `<ul class="suggestions-list">`;
				html += `<li class="urgent">分析下降原因（市场竞争、课程质量、服务问题等）</li>`;
				html += `<li class="urgent">加强营销推广力度</li>`;
				html += `<li>优化课程内容和定价策略</li>`;
				html += `<li>提升客户满意度</li>`;
				html += `</ul></div>`;
			}
		}
		
		// 退费率分析卡片
		const refundRates = quarterData.map(q => ({
			quarter: q.季度,
			rate: q.营收 > 0 ? (q.退费 / q.营收 * 100) : 0
		}));
		const avgRefundRate = refundRates.length > 0 ? refundRates.reduce((sum, r) => sum + r.rate, 0) / refundRates.length : 0;
		const maxRefundRate = refundRates.length > 0 ? Math.max(...refundRates.map(r => r.rate)) : 0;
		const worstRefundQuarter = refundRates.find(r => r.rate === maxRefundRate);
		
		html += `<div class="analysis-section">`;
		html += `<div class="analysis-section-title">💰 退费率</div>`;
		html += `<div class="analysis-text">`;
		html += `<p>平均：<span class="${avgRefundRate < 5 ? 'success' : avgRefundRate < 10 ? 'warning' : 'danger'}">${avgRefundRate.toFixed(1)}%</span></p>`;
		if (worstRefundQuarter && worstRefundQuarter.rate > avgRefundRate * 1.5) {
			html += `<p>最高：<span class="warning">${worstRefundQuarter.quarter}</span> ${worstRefundQuarter.rate.toFixed(1)}%</p>`;
		}
		html += `</div></div>`;
		
		html += `</div>`;
		
		// 添加财务建议
		html += `<div class="insight-box info" style="margin-top: 16px;">`;
		html += `<div class="insight-box-title">💡 财务优化建议</div>`;
		html += `<ul class="suggestions-list">`;
		
		if (avgRefundRate >= 10) {
			html += `<li class="urgent">立即建立退费专项小组，深入分析退费原因，重点关注高退费季度</li>`;
			html += `<li class="urgent">建立退费预警机制，设定退费率"红线"指标（建议<10%）</li>`;
		} else if (avgRefundRate >= 5) {
			html += `<li class="important">持续监控退费率变化，建立退费原因追踪体系</li>`;
		}
		
		if (trendData.length >= 7) {
			const recent = trendData.slice(-7);
			const earlier = trendData.slice(0, Math.min(7, trendData.length - 7));
			const recentAvg = recent.reduce((sum, d) => sum + d.营收, 0) / recent.length;
			const earlierAvg = earlier.length > 0 ? earlier.reduce((sum, d) => sum + d.营收, 0) / earlier.length : recentAvg;
			const trendRate = earlierAvg > 0 ? ((recentAvg - earlierAvg) / earlierAvg * 100) : 0;
			
			if (trendRate <= -10) {
				html += `<li class="urgent">营收下降明显，需立即分析原因并采取应对措施</li>`;
				html += `<li>加强营销推广，优化课程定价策略</li>`;
			} else if (trendRate < 0) {
				html += `<li>关注营收趋势，提前制定应对策略</li>`;
			}
		}
		
		if (quarterData.length > 0) {
			const bestQuarter = quarterData.reduce((max, q) => q.营收 > max.营收 ? q : max, quarterData[0]);
			const worstQuarter = quarterData.reduce((min, q) => q.营收 < min.营收 ? q : min, quarterData[0]);
			const quarterGap = bestQuarter.营收 - worstQuarter.营收;
			const gapRate = worstQuarter.营收 > 0 ? (quarterGap / worstQuarter.营收 * 100) : 0;
			
			if (gapRate > 50) {
				html += `<li>季度营收差异较大，建议制定非旺季的营销策略，平衡季度分布</li>`;
			}
		}
		
		html += `<li>建立完整的财务数据追踪体系，定期进行财务健康度评估</li>`;
		html += `<li>优化成本结构，提高净营收率（目标>70%）</li>`;
		html += `</ul></div>`;
		
		return html;
	}

	// 生成营销分析（优化排版）
	function generateMarketingAnalysis(data) {
		const subjectData = analyzeBySubject(data);
		const classTypeData = analyzeByClassType(data);
		const studentTypeData = analyzeByStudentType(data);
		
		let html = `<div class="analysis-content">`;
		
		// 学科分析卡片
		if (subjectData.length > 0) {
			const topSubject = subjectData[0];
			const totalSubjectRevenue = subjectData.reduce((sum, s) => sum + s.营收, 0);
			const topSubjectShare = totalSubjectRevenue > 0 ? (topSubject.营收 / totalSubjectRevenue * 100) : 0;
			
			html += `<div class="analysis-section">`;
			html += `<div class="analysis-section-title">📚 学科表现</div>`;
			html += `<div class="analysis-text">`;
			html += `<p><span class="success">${topSubject.学科}</span> 领先</p>`;
			html += `<p>${formatToWanYuan(topSubject.营收)} (${topSubjectShare.toFixed(0)}%)</p>`;
			
			if (subjectData.length >= 2) {
				const secondSubject = subjectData[1];
				html += `<p style="margin-top:8px;"><span class="highlight">${secondSubject.学科}</span> 第二</p>`;
				html += `<p>${formatToWanYuan(secondSubject.营收)}</p>`;
				
				// 学科集中度警告
				const top3Share = subjectData.slice(0, 3).reduce((sum, s) => sum + s.营收, 0) / totalSubjectRevenue * 100;
				if (top3Share > 80) {
					html += `<p style="margin-top:8px; color:var(--warn); font-size:11px;">⚠️ 集中度${top3Share.toFixed(0)}%偏高</p>`;
				}
			}
			html += `</div></div>`;
		}
		
		// 班型分析卡片
		if (classTypeData.length > 0) {
			const topClassType = classTypeData[0];
			const avgOrderValue = classTypeData.reduce((sum, ct) => sum + ct.平均订单金额, 0) / classTypeData.length;
			const efficientClassTypes = classTypeData.filter(ct => ct.平均订单金额 > avgOrderValue * 1.2);
			
			html += `<div class="analysis-section">`;
			html += `<div class="analysis-section-title">🎓 班型效率</div>`;
			html += `<div class="analysis-text">`;
			html += `<p><span class="success">${topClassType.班型}</span> 最佳</p>`;
			html += `<p>${formatToWanYuan(topClassType.营收)}</p>`;
			html += `<p style="margin-top:8px; font-size:11px;">均单：${topClassType.平均订单金额.toFixed(0)}元</p>`;
			
			if (efficientClassTypes.length > 0) {
				html += `<p style="margin-top:8px; color:var(--ok); font-size:11px;">✓ 高价值：${efficientClassTypes.map(ct => ct.班型).join('、')}</p>`;
			}
			html += `</div></div>`;
		}
		
		// 学生类型分析卡片
		if (studentTypeData.length > 0) {
			html += `<div class="analysis-section">`;
			html += `<div class="analysis-section-title">👥 学生类型</div>`;
			html += `<div class="analysis-text">`;
			studentTypeData.forEach(st => {
				const statusClass = st.退费率 > 10 ? 'danger' : st.退费率 > 5 ? 'warning' : 'success';
				html += `<p><span class="highlight">${st.学生类型}</span></p>`;
				html += `<p>${formatToWanYuan(st.营收)} | 退费率 <span class="${statusClass}">${st.退费率.toFixed(1)}%</span></p>`;
			});
			html += `</div></div>`;
		}
		
		html += `</div>`;
		
		// 集中度风险提示
		if (subjectData.length > 0) {
			const totalSubjectRevenue = subjectData.reduce((sum, s) => sum + s.营收, 0);
			const top3Share = subjectData.slice(0, 3).reduce((sum, s) => sum + s.营收, 0) / totalSubjectRevenue * 100;
			if (top3Share > 80) {
				html += `<div class="insight-box warning">`;
				html += `<div class="insight-box-title">⚠️ 学科集中度风险</div>`;
				html += `<ul class="suggestions-list">`;
				html += `<li>加强其他学科的营销推广</li>`;
				html += `<li>优化弱势学科的课程内容</li>`;
				html += `<li>平衡学科发展，降低经营风险</li>`;
				html += `</ul></div>`;
			}
		}
		
		// 添加营销建议
		html += `<div class="insight-box info" style="margin-top: 16px;">`;
		html += `<div class="insight-box-title">💡 营销优化建议</div>`;
		html += `<ul class="suggestions-list">`;
		
		if (subjectData.length > 0) {
			const topSubject = subjectData[0];
			const totalSubjectRevenue = subjectData.reduce((sum, s) => sum + s.营收, 0);
			const topSubjectShare = totalSubjectRevenue > 0 ? (topSubject.营收 / totalSubjectRevenue * 100) : 0;
			const topSubjectRefundRate = topSubject.营收 > 0 ? (topSubject.退费 / topSubject.营收 * 100) : 0;
			
			if (topSubjectShare > 50) {
				html += `<li class="important">核心学科（${topSubject.学科}）占比过高，建议：</li>`;
				html += `<li style="padding-left: 20px;">- 加强其他学科的差异化营销</li>`;
				html += `<li style="padding-left: 20px;">- 开发跨学科组合课程</li>`;
			}
			
			if (topSubjectRefundRate > 10) {
				html += `<li class="urgent">核心学科退费率${topSubjectRefundRate.toFixed(1)}%偏高，需重点优化课程质量和服务</li>`;
			}
			
			// 找出表现最好的学科
			const bestSubject = subjectData.reduce((best, s) => {
				const sRefundRate = s.营收 > 0 ? (s.退费 / s.营收 * 100) : 0;
				const bestRefundRate = best.营收 > 0 ? (best.退费 / best.营收 * 100) : 0;
				return sRefundRate < bestRefundRate ? s : best;
			}, subjectData[0]);
			
			if (bestSubject && bestSubject !== topSubject) {
				html += `<li>${bestSubject.学科}表现优异，可作为重点推广学科，复制成功经验</li>`;
			}
		}
		
		if (classTypeData.length > 0) {
			const avgOrderValue = classTypeData.reduce((sum, ct) => sum + ct.平均订单金额, 0) / classTypeData.length;
			const efficientClassTypes = classTypeData.filter(ct => ct.平均订单金额 > avgOrderValue * 1.2);
			
			if (efficientClassTypes.length > 0) {
				html += `<li class="important">高价值班型（${efficientClassTypes.map(ct => ct.班型).join('、')}）表现突出，建议：</li>`;
				html += `<li style="padding-left: 20px;">- 加大营销投入，扩大高价值班型规模</li>`;
				html += `<li style="padding-left: 20px;">- 设计针对性的推广方案</li>`;
			}
			
			const lowValueClassTypes = classTypeData.filter(ct => ct.平均订单金额 < avgOrderValue * 0.8);
			if (lowValueClassTypes.length > 0) {
				html += `<li>低价值班型（${lowValueClassTypes.map(ct => ct.班型).join('、')}）需优化定价或提升价值</li>`;
			}
		}
		
		if (studentTypeData.length > 0) {
			const newStudentData = studentTypeData.find(st => st.学生类型.includes('新生') || st.学生类型.includes('新'));
			const returningStudentData = studentTypeData.find(st => st.学生类型.includes('老生') || st.学生类型.includes('老'));
			
			if (newStudentData && newStudentData.退费率 > 15) {
				html += `<li class="urgent">新生退费率${newStudentData.退费率.toFixed(1)}%偏高，需优化新生体验和课程匹配度</li>`;
			}
			
			if (newStudentData && returningStudentData) {
				const newStudentShare = (newStudentData.营收 / (newStudentData.营收 + returningStudentData.营收) * 100);
				if (newStudentShare > 60) {
					html += `<li>新生占比${newStudentShare.toFixed(0)}%，需加强老生复购和转介绍机制</li>`;
				} else if (newStudentShare < 30) {
					html += `<li>新生占比${newStudentShare.toFixed(0)}%偏低，需加强新客户获取</li>`;
				}
			}
		}
		
		html += `<li>建立完整的营销数据追踪体系，定期分析各维度表现</li>`;
		html += `<li>制定差异化的营销策略，针对不同学科、班型、学生类型优化推广方案</li>`;
		html += `<li>建立客户生命周期管理体系，提升客户留存和复购率</li>`;
		html += `</ul></div>`;
		
		return html;
	}

	// 改进趋势图 - 添加移动平均线和关键点标注
	function drawRevenueTrendChartEnhanced(canvasId, trendData, showMovingAverage = true) {
		const canvas = document.getElementById(canvasId);
		if (!canvas || !trendData || trendData.length === 0) return;
		
		const ctx = canvas.getContext("2d");
		// 清除画布
		ctx.clearRect(0, 0, canvas.width, canvas.height);
		const width = canvas.width;
		const height = canvas.height;
		const padding = { top: 50, right: 40, bottom: 80, left: 80 };
		const chartWidth = width - padding.left - padding.right;
		const chartHeight = height - padding.top - padding.bottom;
		
		ctx.clearRect(0, 0, width, height);
		
		const maxValue = Math.max(...trendData.map(d => Math.max(d.营收, d.退费, d.净营收 || 0)));
		const maxY = Math.ceil(maxValue * 1.15);
		
		// 绘制网格线
		ctx.strokeStyle = "rgba(255, 255, 255, 0.05)";
		ctx.lineWidth = 1;
		for (let i = 0; i <= 5; i++) {
			const y = padding.top + (chartHeight / 5) * i;
			ctx.beginPath();
			ctx.moveTo(padding.left, y);
			ctx.lineTo(padding.left + chartWidth, y);
			ctx.stroke();
		}
		
		// 绘制坐标轴
		ctx.strokeStyle = "#4b5563";
		ctx.lineWidth = 2;
		ctx.beginPath();
		ctx.moveTo(padding.left, padding.top);
		ctx.lineTo(padding.left, padding.top + chartHeight);
		ctx.lineTo(padding.left + chartWidth, padding.top + chartHeight);
		ctx.stroke();
		
		// Y轴刻度
		ctx.fillStyle = "#9ca3af";
		ctx.font = "11px sans-serif";
		ctx.textAlign = "right";
		for (let i = 0; i <= 5; i++) {
			const value = (maxY / 5) * i;
			const y = padding.top + chartHeight - (value / maxY) * chartHeight;
			ctx.fillText(value.toFixed(0), padding.left - 10, y + 4);
		}
		
		// 计算移动平均线（7日）
		const movingAverages = [];
		if (showMovingAverage && trendData.length >= 7) {
			for (let i = 6; i < trendData.length; i++) {
				const window = trendData.slice(i - 6, i + 1);
				const avg = window.reduce((sum, d) => sum + d.营收, 0) / window.length;
				movingAverages.push({ index: i, value: avg });
			}
		}
		
		// 绘制折线
		const pointSpacing = trendData.length > 1 ? chartWidth / (trendData.length - 1) : chartWidth;
		
		// 营收折线
		ctx.beginPath();
		ctx.strokeStyle = "#3b82f6";
		ctx.lineWidth = 3;
		trendData.forEach((d, idx) => {
			const x = padding.left + idx * pointSpacing;
			const y = padding.top + chartHeight - (d.营收 / maxY) * chartHeight;
			if (idx === 0) {
				ctx.moveTo(x, y);
			} else {
				ctx.lineTo(x, y);
			}
		});
		ctx.stroke();
		
		// 绘制数据点
		trendData.forEach((d, idx) => {
			const x = padding.left + idx * pointSpacing;
			const y = padding.top + chartHeight - (d.营收 / maxY) * chartHeight;
			ctx.fillStyle = "#3b82f6";
			ctx.beginPath();
			ctx.arc(x, y, 4, 0, Math.PI * 2);
			ctx.fill();
		});
		
		// 绘制移动平均线
		if (showMovingAverage && movingAverages.length > 0) {
			ctx.beginPath();
			ctx.strokeStyle = "#fbbf24";
			ctx.lineWidth = 2;
			ctx.setLineDash([5, 5]);
			movingAverages.forEach((ma, maIdx) => {
				const x = padding.left + ma.index * pointSpacing;
				const y = padding.top + chartHeight - (ma.value / maxY) * chartHeight;
				if (maIdx === 0) {
					ctx.moveTo(x, y);
				} else {
					ctx.lineTo(x, y);
				}
			});
			ctx.stroke();
			ctx.setLineDash([]);
		}
		
		// 标注最高点和最低点
		const maxPoint = trendData.reduce((max, d, idx) => d.营收 > max.revenue ? { revenue: d.营收, index: idx, date: d.日期 } : max, { revenue: 0, index: 0, date: '' });
		const minPoint = trendData.reduce((min, d, idx) => d.营收 < min.revenue ? { revenue: d.营收, index: idx, date: d.日期 } : min, { revenue: Infinity, index: 0, date: '' });
		
		// 标注最高点
		if (maxPoint.revenue > 0) {
			const x = padding.left + maxPoint.index * pointSpacing;
			const y = padding.top + chartHeight - (maxPoint.revenue / maxY) * chartHeight;
			ctx.fillStyle = "#10b981";
			ctx.beginPath();
			ctx.arc(x, y, 6, 0, Math.PI * 2);
			ctx.fill();
			ctx.fillStyle = "#fff";
			ctx.font = "bold 10px sans-serif";
			ctx.textAlign = "center";
			ctx.fillText("最高", x, y - 10);
		}
		
		// 标注最低点
		if (minPoint.revenue < Infinity) {
			const x = padding.left + minPoint.index * pointSpacing;
			const y = padding.top + chartHeight - (minPoint.revenue / maxY) * chartHeight;
			ctx.fillStyle = "#ef4444";
			ctx.beginPath();
			ctx.arc(x, y, 6, 0, Math.PI * 2);
			ctx.fill();
			ctx.fillStyle = "#fff";
			ctx.font = "bold 10px sans-serif";
			ctx.textAlign = "center";
			ctx.fillText("最低", x, y + 20);
		}
		
		// X轴标签（日期）
		ctx.fillStyle = "#9ca3af";
		ctx.font = "10px sans-serif";
		ctx.textAlign = "center";
		const labelStep = Math.max(1, Math.floor(trendData.length / 10));
		trendData.forEach((d, idx) => {
			if (idx % labelStep === 0 || idx === trendData.length - 1) {
				const x = padding.left + idx * pointSpacing;
				const date = new Date(d.日期);
				const label = `${date.getMonth() + 1}/${date.getDate()}`;
				ctx.save();
				ctx.translate(x, padding.top + chartHeight + 15);
				ctx.rotate(-Math.PI / 4);
				ctx.fillText(label, 0, 0);
				ctx.restore();
			}
		});
		
		// 图例
		ctx.fillStyle = "#3b82f6";
		ctx.fillRect(padding.left + chartWidth - 120, padding.top - 30, 12, 12);
		ctx.fillStyle = "#e5e7eb";
		ctx.font = "11px sans-serif";
		ctx.textAlign = "left";
		ctx.fillText("营收", padding.left + chartWidth - 105, padding.top - 20);
		
		if (showMovingAverage && movingAverages.length > 0) {
			ctx.strokeStyle = "#fbbf24";
			ctx.lineWidth = 2;
			ctx.setLineDash([5, 5]);
			ctx.beginPath();
			ctx.moveTo(padding.left + chartWidth - 120, padding.top - 15);
			ctx.lineTo(padding.left + chartWidth - 100, padding.top - 15);
			ctx.stroke();
			ctx.setLineDash([]);
			ctx.fillText("7日移动平均", padding.left + chartWidth - 95, padding.top - 10);
		}
	}

	// 生成图表洞察
	function generateChartInsights(chartType, data) {
		let insights = [];
		
		if (chartType === "revenueTrend") {
			if (data.length >= 7) {
				const recent = data.slice(-7);
				const earlier = data.slice(0, Math.min(7, data.length - 7));
				const recentAvg = recent.reduce((sum, d) => sum + d.营收, 0) / recent.length;
				const earlierAvg = earlier.length > 0 ? earlier.reduce((sum, d) => sum + d.营收, 0) / earlier.length : recentAvg;
				const trend = recentAvg - earlierAvg;
				const trendRate = earlierAvg > 0 ? (trend / earlierAvg * 100) : 0;
				
				if (Math.abs(trendRate) > 5) {
					insights.push(`近期营收${trendRate > 0 ? '上升' : '下降'} ${Math.abs(trendRate).toFixed(1)}%`);
				}
				
				const maxPoint = data.reduce((max, d) => d.营收 > max ? d.营收 : max, 0);
				const minPoint = data.reduce((min, d) => d.营收 < min ? d.营收 : min, Infinity);
				if (maxPoint > 0 && minPoint < Infinity) {
					insights.push(`最高点与最低点相差 ${((maxPoint - minPoint) / minPoint * 100).toFixed(1)}%`);
				}
			}
		}
		
		return insights;
	}

	// AI分析功能（可选，需要用户明确启用）
	async function generateAIAnalysis(data, apiKey, provider = "deepseek") {
		if (provider === "local") {
			return generateLocalAnalysis(data);
		}
		
		// 数据脱敏
		const sanitizedData = sanitizeDataForAI(data);
		
		try {
			const response = await fetch("https://api.deepseek.com/v1/chat/completions", {
				method: "POST",
				headers: {
					"Content-Type": "application/json",
					"Authorization": `Bearer ${apiKey}`
				},
				body: JSON.stringify({
					model: "deepseek-chat",
					messages: [{
						role: "system",
						content: "你是一位专业的教育培训机构数据分析师，擅长从财务和营销角度分析业务数据，提供 actionable insights。"
					}, {
						role: "user",
						content: `请基于以下数据生成详细的分析报告，包括：1. 总体情况分析 2. 财务健康度评估 3. 营销效果分析 4. 风险提示 5. 改进建议。数据：${JSON.stringify(sanitizedData)}`
					}],
					temperature: 0.7,
					max_tokens: 2000
				})
			});
			
			if (!response.ok) {
				throw new Error(`API请求失败: ${response.statusText}`);
			}
			
			const result = await response.json();
			return {
				overview: result.choices[0].message.content,
				financial: result.choices[0].message.content,
				marketing: result.choices[0].message.content
			};
		} catch (error) {
			console.error("AI分析失败:", error);
			throw error;
		}
	}

	// ==================== 图表筛选功能 ====================
	
	let chartFilters = {
		quarter: new Set(),
		subject: new Set(),
		classType: new Set(),
		studentType: new Set()
	};
	
	// 根据筛选条件过滤数据
	function filterDataByChartFilters(data, filters) {
		// 确保filters对象存在且包含必要的属性
		if (!filters) {
			console.warn("filterDataByChartFilters: filters参数为空，返回原始数据");
			return data;
		}
		
		let filteredRevenue = [...data.revenue];
		let filteredRefund = [...data.refund];
		
		// 按季度筛选
		if (filters.quarter && filters.quarter.size > 0) {
			filteredRevenue = filteredRevenue.filter(r => {
				const quarter = r["季度"] || "";
				return filters.quarter.has(quarter);
			});
			filteredRefund = filteredRefund.filter(r => {
				const quarter = r["季度"] || r["学期"] || "";
				return filters.quarter.has(quarter);
			});
		}
		
		// 按学科筛选
		if (filters.subject && filters.subject.size > 0) {
			filteredRevenue = filteredRevenue.filter(r => {
				const subject = r["学科"] || "";
				return filters.subject.has(subject);
			});
			filteredRefund = filteredRefund.filter(r => {
				const subject = r["科目"] || "";
				return filters.subject.has(subject);
			});
		}
		
		// 按班型筛选
		if (filters.classType && filters.classType.size > 0) {
			filteredRevenue = filteredRevenue.filter(r => {
				const classType = r["班型"] || "";
				return filters.classType.has(classType);
			});
		}
		
		// 按学生类型筛选
		if (filters.studentType && filters.studentType.size > 0) {
			filteredRevenue = filteredRevenue.filter(r => {
				const studentType = r["学生类型"] || "";
				return filters.studentType.has(studentType);
			});
			filteredRefund = filteredRefund.filter(r => {
				const studentType = r["是否新生"] === "是" ? "新生" : (r["是否新生"] === "否" ? "老生" : "");
				return filters.studentType.has(studentType);
			});
		}
		
		return { revenue: filteredRevenue, refund: filteredRefund };
	}
	
	// 绑定筛选选项事件（在选项生成后调用）
	function bindFilterOptionEvents() {
		// 使用事件委托，避免重复绑定问题
		const financialTab = document.getElementById("financialTab");
		const marketingTab = document.getElementById("marketingTab");
		
		// 财务分析筛选选项 - 使用事件委托
		if (financialTab) {
			financialTab.addEventListener("change", (e) => {
				if (e.target.classList.contains("filter-option")) {
					// 检查是否有筛选维度被启用
					const hasActiveFilter = Array.from(document.querySelectorAll("#financialTab .chart-filter-checkbox")).some(cb => cb.checked);
					if (hasActiveFilter) {
						// 使用防抖，避免频繁渲染
						if (typeof debouncedApplyFinancialFilter !== 'undefined') {
							debouncedApplyFinancialFilter();
						} else {
							setTimeout(() => applyFinancialChartFilter(), 300);
						}
					} else {
						// 如果没有筛选维度被启用，重置为全部数据
						const data = getData();
						const dateFiltered = filterByDateRange(data, 
							document.getElementById("startDate")?.value,
							document.getElementById("endDate")?.value
						);
						renderAllAnalysis(dateFiltered);
						const localAnalysis = generateFinancialAnalysis(dateFiltered);
						document.getElementById("financialAnalysis").innerHTML = localAnalysis;
					}
				}
			});
		}
		
		// 营销分析筛选选项 - 使用事件委托
		if (marketingTab) {
			marketingTab.addEventListener("change", (e) => {
				if (e.target.classList.contains("filter-option")) {
					// 检查是否有筛选维度被启用
					const hasActiveFilter = Array.from(document.querySelectorAll("#marketingTab .chart-filter-checkbox")).some(cb => cb.checked);
					if (hasActiveFilter) {
						debouncedApplyMarketingFilter();
					} else {
						// 如果没有筛选维度被启用，重置为全部数据
						const data = getData();
						const dateFiltered = filterByDateRange(data, 
							document.getElementById("startDate")?.value,
							document.getElementById("endDate")?.value
						);
						renderAllAnalysis(dateFiltered);
						const localAnalysis = generateMarketingAnalysis(dateFiltered);
						document.getElementById("marketingAnalysis").innerHTML = localAnalysis;
					}
				}
			});
		}
	}
	
	// 初始化筛选选项
	function initChartFilters(data) {
		// 季度选项
		const quarters = new Set();
		data.revenue.forEach(r => {
			const q = r["季度"] || "";
			if (q) quarters.add(q);
		});
		data.refund.forEach(r => {
			const q = r["季度"] || r["学期"] || "";
			if (q) quarters.add(q);
		});
		
		const quarterContainer = document.getElementById("quarterCheckboxes");
		if (quarterContainer) {
			quarterContainer.innerHTML = Array.from(quarters).sort().map(q => 
				`<label><input type="checkbox" value="${q}" class="filter-option"> <span>${q}</span></label>`
			).join('');
		}
		
		// 学科选项
		const subjects = new Set();
		data.revenue.forEach(r => {
			const s = r["学科"] || "";
			if (s) subjects.add(s);
		});
		data.refund.forEach(r => {
			const s = r["科目"] || "";
			if (s) subjects.add(s);
		});
		
		const subjectContainer = document.getElementById("subjectCheckboxes");
		const subjectMarketingContainer = document.getElementById("subjectMarketingCheckboxes");
		const subjectOptions = Array.from(subjects).sort().map(s => 
			`<label><input type="checkbox" value="${s}" class="filter-option"> <span>${s}</span></label>`
		).join('');
		
		if (subjectContainer) subjectContainer.innerHTML = subjectOptions;
		if (subjectMarketingContainer) subjectMarketingContainer.innerHTML = subjectOptions;
		
		// 班型选项
		const classTypes = new Set();
		data.revenue.forEach(r => {
			const ct = r["班型"] || "";
			if (ct) classTypes.add(ct);
		});
		
		const classTypeContainer = document.getElementById("classTypeCheckboxes");
		if (classTypeContainer) {
			classTypeContainer.innerHTML = Array.from(classTypes).sort().map(ct => 
				`<label><input type="checkbox" value="${ct}" class="filter-option"> <span>${ct}</span></label>`
			).join('');
		}
		
		// 学生类型选项
		const studentTypes = new Set();
		data.revenue.forEach(r => {
			const st = r["学生类型"] || "";
			if (st) studentTypes.add(st);
		});
		data.refund.forEach(r => {
			const st = r["是否新生"] === "是" ? "新生" : (r["是否新生"] === "否" ? "老生" : "");
			if (st) studentTypes.add(st);
		});
		
		const studentTypeContainer = document.getElementById("studentTypeCheckboxes");
		if (studentTypeContainer) {
			studentTypeContainer.innerHTML = Array.from(studentTypes).sort().map(st => 
				`<label><input type="checkbox" value="${st}" class="filter-option"> <span>${st}</span></label>`
			).join('');
		}
		
		// 绑定筛选选项事件
		setTimeout(() => {
			bindFilterOptionEvents();
		}, 100);
	}
	
	// 应用财务图表筛选（优化版本：先更新图表，后更新报告）
	function applyFinancialChartFilter() {
		// 如果正在执行，取消之前的操作
		if (isFiltering) {
			console.log("取消之前的筛选操作");
			if (debouncedApplyFinancialFilter && debouncedApplyFinancialFilter.cancel) {
				debouncedApplyFinancialFilter.cancel();
			}
			pendingFilterCancel = true;
		}
		
		// 设置执行状态
		isFiltering = true;
		pendingFilterCancel = false;
		
		console.log("开始应用财务筛选...");
		
		// 显示加载状态
		const applyBtn = document.getElementById("applyChartFilter");
		const originalText = applyBtn?.textContent || "应用筛选";
		if (applyBtn) {
			applyBtn.textContent = "处理中...";
			applyBtn.disabled = true;
		}
		
		// 使用requestAnimationFrame确保UI响应，并支持取消
		const rafId = requestAnimationFrame(() => {
			// 检查是否被取消
			if (pendingFilterCancel) {
				console.log("筛选操作已取消");
				isFiltering = false;
				if (applyBtn) {
					applyBtn.textContent = originalText;
					applyBtn.disabled = false;
				}
				return;
			}
			
			try {
				const filters = {
					quarter: new Set(),
					subject: new Set()
				};
				
				// 收集季度筛选
				if (document.getElementById("filterByQuarter")?.checked) {
					document.querySelectorAll("#quarterCheckboxes input:checked").forEach(cb => {
						filters.quarter.add(cb.value);
					});
				}
				
				// 收集学科筛选
				if (document.getElementById("filterBySubject")?.checked) {
					document.querySelectorAll("#subjectCheckboxes input:checked").forEach(cb => {
						filters.subject.add(cb.value);
					});
				}
				
				console.log("筛选条件:", {
					季度: Array.from(filters.quarter),
					学科: Array.from(filters.subject)
				});
				
				const data = getData();
				const dateFiltered = filterByDateRange(data, 
					document.getElementById("startDate")?.value,
					document.getElementById("endDate")?.value
				);
				const filteredData = filterDataByChartFilters(dateFiltered, filters);
				
				console.log("筛选后数据量:", {
					营收: filteredData.revenue.length,
					退费: filteredData.refund.length
				});
				
				// 优先更新图表（同步执行，确保立即显示）
				const quarterData = analyzeByQuarter(filteredData);
				const trendData = analyzeRevenueTrend(filteredData);
				const showMovingAvg = document.getElementById("showMovingAverage")?.checked ?? true;
				
				renderTableData("quarterRevenueTable", quarterData);
				if (trendData.length > 0) {
					// 确保图表正确更新，包括移动平均线选项
					// 强制重新绘制，清除之前的图表
					const revenueCanvas = document.getElementById("revenueTrendChart");
					const netCanvas = document.getElementById("netRevenueTrendChart");
					if (revenueCanvas) {
						const ctx = revenueCanvas.getContext("2d");
						ctx.clearRect(0, 0, revenueCanvas.width, revenueCanvas.height);
					}
					if (netCanvas) {
						const ctx = netCanvas.getContext("2d");
						ctx.clearRect(0, 0, netCanvas.width, netCanvas.height);
					}
					
					drawRevenueTrendChartEnhanced("revenueTrendChart", trendData, showMovingAvg);
					drawRevenueTrendChartEnhanced("netRevenueTrendChart", trendData, showMovingAvg);
					drawRefundRateChart("refundRateChart", quarterData);
					console.log("图表已更新");
				}
				
				// 恢复按钮状态
				isFiltering = false;
				if (applyBtn) {
					applyBtn.textContent = originalText;
					applyBtn.disabled = false;
				}
				
				// 延迟更新分析报告（异步执行，不阻塞图表显示）
				setTimeout(() => {
					// 再次检查是否被取消
					if (pendingFilterCancel) {
						return;
					}
					const localAnalysis = generateFinancialAnalysis(filteredData);
					const analysisEl = document.getElementById("financialAnalysis");
					if (analysisEl) {
						analysisEl.innerHTML = localAnalysis;
						console.log("分析报告已更新");
					}
				}, 50);
			} catch (error) {
				console.error("应用财务筛选时出错:", error);
				isFiltering = false;
				if (applyBtn) {
					applyBtn.textContent = originalText;
					applyBtn.disabled = false;
				}
			}
		});
		
		// 保存rafId以便取消（如果需要）
		return rafId;
	}
	
	// 应用营销图表筛选（优化版本：先更新图表，后更新报告）
	function applyMarketingChartFilter() {
		// 如果正在执行，取消之前的操作
		if (isFiltering) {
			console.log("取消之前的筛选操作");
			if (debouncedApplyMarketingFilter && debouncedApplyMarketingFilter.cancel) {
				debouncedApplyMarketingFilter.cancel();
			}
			pendingFilterCancel = true;
		}
		
		// 设置执行状态
		isFiltering = true;
		pendingFilterCancel = false;
		
		console.log("开始应用营销筛选...");
		
		// 显示加载状态
		const applyBtn = document.getElementById("applyMarketingFilter");
		const originalText = applyBtn?.textContent || "应用筛选";
		if (applyBtn) {
			applyBtn.textContent = "处理中...";
			applyBtn.disabled = true;
		}
		
		// 使用requestAnimationFrame确保UI响应，并支持取消
		const rafId = requestAnimationFrame(() => {
			// 检查是否被取消
			if (pendingFilterCancel) {
				console.log("筛选操作已取消");
				isFiltering = false;
				if (applyBtn) {
					applyBtn.textContent = originalText;
					applyBtn.disabled = false;
				}
				return;
			}
			
			try {
				const filters = {
					subject: new Set(),
					classType: new Set(),
					studentType: new Set()
				};
				
				// 收集筛选条件
				if (document.getElementById("filterBySubjectMarketing")?.checked) {
					document.querySelectorAll("#subjectMarketingCheckboxes input:checked").forEach(cb => {
						filters.subject.add(cb.value);
					});
				}
				
				if (document.getElementById("filterByClassType")?.checked) {
					document.querySelectorAll("#classTypeCheckboxes input:checked").forEach(cb => {
						filters.classType.add(cb.value);
					});
				}
				
				if (document.getElementById("filterByStudentType")?.checked) {
					document.querySelectorAll("#studentTypeCheckboxes input:checked").forEach(cb => {
						filters.studentType.add(cb.value);
					});
				}
				
				console.log("筛选条件:", {
					学科: Array.from(filters.subject),
					班型: Array.from(filters.classType),
					学生类型: Array.from(filters.studentType)
				});
				
				const data = getData();
				const dateFiltered = filterByDateRange(data, 
					document.getElementById("startDate")?.value,
					document.getElementById("endDate")?.value
				);
				const filteredData = filterDataByChartFilters(dateFiltered, filters);
				
				console.log("筛选后数据量:", {
					营收: filteredData.revenue.length,
					退费: filteredData.refund.length
				});
				
				// 优先更新图表（同步执行，确保立即显示）
				const subjectData = analyzeBySubject(filteredData);
				if (subjectData.length > 0) {
					drawPieChart("subjectRevenueChart", subjectData.map(s => ({
						销售人员: s.学科,
						营收: s.营收,
						退费金额: s.退费,
						净营收: s.净营收
					})));
				}
				
				const subjectPeriodData = analyzeSubjectByPeriod(filteredData);
				renderTableData("subjectPeriodMatrix", subjectPeriodData.data, ["学科", ...subjectPeriodData.periods]);
				
				// 绘制学科×周期柱状图
				if (subjectPeriodData.data.length > 0 && subjectPeriodData.periods.length > 0) {
					const canvas = document.getElementById("subjectPeriodChart");
					if (canvas) {
						canvas.width = 720;
					}
					drawSubjectPeriodChart("subjectPeriodChart", subjectPeriodData);
				}
				
				const classTypeData = analyzeByClassType(filteredData);
				if (classTypeData.length > 0) {
					drawPieChart("classTypeChart", classTypeData.map(ct => ({
						销售人员: ct.班型,
						营收: ct.营收,
						退费金额: 0,
						净营收: ct.营收
					})));
				}
				
				const studentTypeData = analyzeByStudentType(filteredData);
				if (studentTypeData.length > 0) {
					drawPieChart("studentTypeChart", studentTypeData.map(st => ({
						销售人员: st.学生类型,
						营收: st.营收,
						退费金额: st.退费,
						净营收: st.净营收
					})));
				}
				
				console.log("图表已更新");
				
				// 恢复按钮状态
				isFiltering = false;
				if (applyBtn) {
					applyBtn.textContent = originalText;
					applyBtn.disabled = false;
				}
				
				// 延迟更新分析报告（异步执行，不阻塞图表显示）
				setTimeout(() => {
					// 再次检查是否被取消
					if (pendingFilterCancel) {
						return;
					}
					const localAnalysis = generateMarketingAnalysis(filteredData);
					const analysisEl = document.getElementById("marketingAnalysis");
					if (analysisEl) {
						analysisEl.innerHTML = localAnalysis;
						console.log("分析报告已更新");
					}
				}, 50);
			} catch (error) {
				console.error("应用营销筛选时出错:", error);
				isFiltering = false;
				if (applyBtn) {
					applyBtn.textContent = originalText;
					applyBtn.disabled = false;
				}
			}
		});
		
		// 保存rafId以便取消（如果需要）
		return rafId;
	}

	// 创建防抖版本的筛选函数（延迟定义，因为函数在DOMContentLoaded之后才可用）
	let debouncedApplyFinancialFilter, debouncedApplyMarketingFilter;
	
	document.addEventListener("DOMContentLoaded", () => {
		// 在函数定义后创建防抖版本（增加延迟，减少频繁触发）
		debouncedApplyFinancialFilter = debounce(applyFinancialChartFilter, 300);
		debouncedApplyMarketingFilter = debounce(applyMarketingChartFilter, 300);
		main();
		initTabs();
		
		// 初始化筛选选项（延迟执行，确保数据已加载）
		setTimeout(() => {
			const data = getData();
			if (data.revenue.length > 0 || data.refund.length > 0) {
				initChartFilters(data);
			}
		}, 200);
		
		// 财务分析筛选器 - 添加实时更新功能
		document.querySelectorAll("#financialTab .chart-filter-checkbox").forEach(checkbox => {
			checkbox.addEventListener("change", (e) => {
				const filterType = e.target.getAttribute("data-filter");
				let optionsDiv = null;
				if (filterType === "quarter") {
					optionsDiv = document.getElementById("quarterFilterOptions");
				} else if (filterType === "subject") {
					optionsDiv = document.getElementById("subjectFilterOptions");
				}
				if (optionsDiv) {
					optionsDiv.style.display = e.target.checked ? "block" : "none";
				}
				// 检查是否有任何筛选器被启用
				const hasActiveFilter = Array.from(document.querySelectorAll("#financialTab .chart-filter-checkbox")).some(cb => cb.checked);
				const filterOptionsDiv = document.getElementById("filterOptions");
				if (filterOptionsDiv) {
					filterOptionsDiv.style.display = hasActiveFilter ? "block" : "none";
				}
				
				// 实时更新图表（使用防抖，避免频繁渲染）
				// 注意：即使没有选择具体选项，也要更新（显示全部数据）
				if (typeof debouncedApplyFinancialFilter !== 'undefined') {
					debouncedApplyFinancialFilter();
				} else {
					setTimeout(() => applyFinancialChartFilter(), 300);
				}
			});
		});
		
		// 筛选选项事件已在initChartFilters中绑定，这里不需要重复绑定
		
		// 绑定财务分析的"应用筛选"按钮
		document.getElementById("applyChartFilter")?.addEventListener("click", () => {
			applyFinancialChartFilter();
		});
		
		// 绑定时间筛选的"应用筛选"按钮（同时应用到当前激活的标签页）
		document.getElementById("applyFilterBtn")?.addEventListener("click", () => {
			const activeTab = document.querySelector(".tab-btn.active")?.getAttribute("data-tab");
			if (activeTab === "financial") {
				applyFinancialChartFilter();
			} else if (activeTab === "marketing") {
				applyMarketingChartFilter();
			} else {
				// 概览或销售分析，使用renderAllAnalysis
				const data = getData();
				const filteredData = filterByDateRange(data, 
					document.getElementById("startDate")?.value,
					document.getElementById("endDate")?.value
				);
				renderAllAnalysis(filteredData);
			}
		});
		
		// 绑定时间快捷选择按钮
		document.querySelectorAll(".date-quick-btn").forEach(btn => {
			btn.addEventListener("click", () => {
				const days = parseInt(btn.getAttribute("data-days"));
				const endDate = new Date();
				const startDate = new Date();
				startDate.setDate(startDate.getDate() - days);
				
				const startDateInput = document.getElementById("startDate");
				const endDateInput = document.getElementById("endDate");
				if (startDateInput) {
					startDateInput.value = startDate.toISOString().split('T')[0];
				}
				if (endDateInput) {
					endDateInput.value = endDate.toISOString().split('T')[0];
				}
				
				// 自动应用筛选
				const activeTab = document.querySelector(".tab-btn.active")?.getAttribute("data-tab");
				if (activeTab === "financial") {
					applyFinancialChartFilter();
				} else if (activeTab === "marketing") {
					applyMarketingChartFilter();
				} else {
					const data = getData();
					const filteredData = filterByDateRange(data, 
						startDateInput?.value,
						endDateInput?.value
					);
					renderAllAnalysis(filteredData);
				}
			});
		});
		document.getElementById("resetChartFilter")?.addEventListener("click", () => {
			// 重置筛选
			document.querySelectorAll("#financialTab .chart-filter-checkbox").forEach(cb => cb.checked = false);
			document.querySelectorAll("#financialTab .filter-option").forEach(cb => cb.checked = false);
			document.getElementById("filterOptions").style.display = "none";
			
			// 重新渲染
			const data = getData();
			const filteredData = filterByDateRange(data, 
				document.getElementById("startDate")?.value,
				document.getElementById("endDate")?.value
			);
			renderAllAnalysis(filteredData);
			const localAnalysis = generateFinancialAnalysis(filteredData);
			document.getElementById("financialAnalysis").innerHTML = localAnalysis;
		});
		
		// 营销分析筛选器 - 添加实时更新功能
		document.querySelectorAll("#marketingTab .chart-filter-checkbox").forEach(checkbox => {
			checkbox.addEventListener("change", (e) => {
				const filterType = e.target.getAttribute("data-filter");
				let optionsDiv = null;
				if (filterType === "subject") {
					optionsDiv = document.getElementById("subjectMarketingFilterOptions");
				} else if (filterType === "classType") {
					optionsDiv = document.getElementById("classTypeFilterOptions");
				} else if (filterType === "studentType") {
					optionsDiv = document.getElementById("studentTypeFilterOptions");
				}
				if (optionsDiv) {
					optionsDiv.style.display = e.target.checked ? "block" : "none";
				}
				// 检查是否有任何筛选器被启用
				const hasActiveFilter = Array.from(document.querySelectorAll("#marketingTab .chart-filter-checkbox")).some(cb => cb.checked);
				const filterOptionsDiv = document.getElementById("marketingFilterOptions");
				if (filterOptionsDiv) {
					filterOptionsDiv.style.display = hasActiveFilter ? "block" : "none";
				}
				
				// 实时更新图表（使用防抖，避免频繁渲染）
				// 注意：即使没有选择具体选项，也要更新（显示全部数据）
				if (typeof debouncedApplyMarketingFilter !== 'undefined') {
					debouncedApplyMarketingFilter();
				} else {
					setTimeout(() => applyMarketingChartFilter(), 300);
				}
			});
		});
		
		// 筛选选项事件已在initChartFilters中绑定，这里不需要重复绑定
		
		document.getElementById("applyMarketingFilter")?.addEventListener("click", applyMarketingChartFilter);
		document.getElementById("resetMarketingFilter")?.addEventListener("click", () => {
			// 重置筛选
			document.querySelectorAll("#marketingTab .chart-filter-checkbox").forEach(cb => cb.checked = false);
			document.querySelectorAll("#marketingTab .filter-option").forEach(cb => cb.checked = false);
			document.getElementById("marketingFilterOptions").style.display = "none";
			
			// 重新渲染（使用applyMarketingChartFilter的逻辑）
			const data = getData();
			const dateFiltered = filterByDateRange(data, 
				document.getElementById("startDate")?.value,
				document.getElementById("endDate")?.value
			);
			
			// 重新渲染所有营销分析图表
			const subjectData = analyzeBySubject(dateFiltered);
			drawPieChart("subjectRevenueChart", subjectData.map(s => ({
				销售人员: s.学科,
				营收: s.营收,
				退费金额: s.退费,
				净营收: s.净营收
			})));
			
			const subjectPeriodData = analyzeSubjectByPeriod(dateFiltered);
			renderTableData("subjectPeriodMatrix", subjectPeriodData.data, ["学科", ...subjectPeriodData.periods]);
			
			if (subjectPeriodData.data.length > 0 && subjectPeriodData.periods.length > 0) {
				const canvas = document.getElementById("subjectPeriodChart");
				if (canvas) {
					canvas.width = 600 + 120;
				}
				drawSubjectPeriodChart("subjectPeriodChart", subjectPeriodData);
			}
			
			const classTypeData = analyzeByClassType(dateFiltered);
			drawPieChart("classTypeChart", classTypeData.map(ct => ({
				销售人员: ct.班型,
				营收: ct.营收,
				退费金额: 0,
				净营收: ct.营收
			})));
			
			const studentTypeData = analyzeByStudentType(dateFiltered);
			drawPieChart("studentTypeChart", studentTypeData.map(st => ({
				销售人员: st.学生类型,
				营收: st.营收,
				退费金额: st.退费,
				净营收: st.净营收
			})));
			
			const localAnalysis = generateMarketingAnalysis(dateFiltered);
			document.getElementById("marketingAnalysis").innerHTML = localAnalysis;
		});
		
		// AI分析配置
		const enableAICheckbox = document.getElementById("enableAIAnalysis");
		const aiConfig = document.getElementById("aiConfig");
		const generateAIBtn = document.getElementById("generateAIAnalysis");
		
		if (enableAICheckbox) {
			enableAICheckbox.addEventListener("change", (e) => {
				aiConfig.style.display = e.target.checked ? "block" : "none";
			});
		}
		
		if (generateAIBtn) {
			generateAIBtn.addEventListener("click", async () => {
				const provider = document.getElementById("aiProvider")?.value || "local";
				const apiKey = document.getElementById("aiApiKey")?.value || "";
				
				if (provider === "deepseek" && !apiKey) {
					alert("请输入DeepSeek API密钥");
					return;
				}
				
				generateAIBtn.disabled = true;
				generateAIBtn.textContent = "分析中...";
				
				try {
					const data = getData();
					const analysis = provider === "local" 
						? generateLocalAnalysis(data)
						: await generateAIAnalysis(data, apiKey, provider);
					
					// 更新分析报告
					document.getElementById("overviewAnalysis").innerHTML = analysis.overview;
					document.getElementById("financialAnalysis").innerHTML = analysis.financial;
					document.getElementById("marketingAnalysis").innerHTML = analysis.marketing;
					
					alert("分析报告已生成！");
				} catch (error) {
					alert("生成分析报告失败：" + error.message);
				} finally {
					generateAIBtn.disabled = false;
					generateAIBtn.textContent = "生成AI分析报告";
				}
			});
		}
		
		// 移动平均线切换
		const movingAvgCheckbox = document.getElementById("showMovingAverage");
		if (movingAvgCheckbox) {
			movingAvgCheckbox.addEventListener("change", (e) => {
				applyFinancialChartFilter();
			});
		}
	});
})();



