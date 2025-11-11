$(document).ready(function() {
    console.log('Query.js loaded successfully!');
    
    // 全局变量
    let currentQueryOrderNumber = '';
    let currentQueryPage = 1;
    let currentQueryPageSize = 10;
    let allMachineTypes = [];
    let currentMachineType = null;

    // 初始化页面
    console.log('Initializing query page...');
    initializeQueryPage();
    
    // 初始化用户信息显示
    initializeUserDisplay();
    
    // 登出按钮事件 - 先移除之前的事件绑定，避免重复绑定
    $('#logoutBtn').off('click').on('click', function() {
        performLogout();
    });

    // 查询按钮点击事件
    $('#queryBtn').click(function() {
        const orderNumber = $('#orderNumberQuery').val().trim();
        if (!orderNumber) {
            showAlert('请输入送检单号', 'error');
            return;
        }
        
        currentQueryOrderNumber = orderNumber;
        currentQueryPage = 1;
        queryOrderByNumber(orderNumber, 1);
    });

    // 机台类型选择
    $('#machineGroupSelector').change(function() {
        const selectedType = $(this).val();
        if (selectedType && selectedType !== '') {
            currentMachineType = selectedType;
            loadMachineStatusByType(selectedType);
        }
    });

    // 显示所有按钮
    $('#showAllBtn').click(function() {
        currentMachineType = null;
        loadAllMachineStatus();
    });

    // 分页按钮事件
    $('#prevPageQuery').click(function() {
        if (currentQueryPage > 1) {
            currentQueryPage--;
            queryOrderByNumber(currentQueryOrderNumber, currentQueryPage);
        }
    });

    $('#nextPageQuery').click(function() {
        queryOrderByNumber(currentQueryOrderNumber, currentQueryPage + 1);
    });

    // 关闭机台作业列表
    $('#closeMachineList').click(function() {
        $('#machineOperationSection').hide();
        $('#queryInfoSection').hide();
    });

    // 初始化页面
    function initializeQueryPage() {
        console.log('Starting page initialization...');
        
        // 先加载所有机台类型，然后加载机台状态
        console.log('Loading machine types...');
        loadAllMachineTypes();
        
        console.log('Page initialization completed!');
    }

    // 加载所有机台类型
    function loadAllMachineTypes() {
        $.ajax({
            url: '/QmsSchedule/GetAllMachineTypes',
            type: 'POST',
            dataType: 'json',
            success: function(response) {
                if (response.status === 'success') {
                    // 仅保留前端允许的三种类型：QV、三次元、UA3P
                    var rawTypes = response.result || [];
                    var filtered = rawTypes.filter(function(t) {
                        var name = (t.QcMachineModeName || '') + ' ' + (t.QcMachineModeDesc || '') + ' ' + (t.QcMachineModeNo || '');
                        return /\bQV\b/i.test(name) || /UA3P/i.test(name) || name.indexOf('三次元') >= 0;
                    });
                    allMachineTypes = filtered;
                    console.log('机台类型加载成功(已过滤到 QV/三次元/UA3P):', allMachineTypes);
                    updateMachineTypeSelector();
                    
                    // 机台类型加载成功后，加载所有机台状态
                    console.log('Loading all machine status...');
                    loadAllMachineStatus();
                } else {
                    console.error('获取机台类型失败:', response.msg);
                }
            },
            error: function(xhr, status, error) {
                console.error('获取机台类型API调用失败:', error);
            }
        });
    }

    // 更新机台类型选择器
    function updateMachineTypeSelector() {
        const selector = $('#machineGroupSelector');
        selector.empty();
        selector.append('<option value="">请选择机台类型</option>');
        
        allMachineTypes.forEach(function(type) {
            const optionText = `${type.QcMachineModeName} (${type.QcMachineModeNo})`;
            selector.append(`<option value="${type.QcMachineModeId}">${optionText}</option>`);
        });
        
        // 生成机台状态区域
        generateMachineStatusSections();
    }

    // 生成机台状态区域
    function generateMachineStatusSections() {
        const container = $('#machineStatusSections');
        container.empty();
        
        allMachineTypes.forEach(function(type) {
            const sectionHtml = `
                <div class="col-12 mb-4" id="machineStatusSection${type.QcMachineModeId}">
                    <div class="card">
                        <div class="card-header">
                            <h4 class="mb-0">${type.QcMachineModeName} 量测机台状态</h4>
                        </div>
                        <div class="card-body">
                            <div class="row" id="machineStatusCards${type.QcMachineModeId}">
                                <!-- 机台状态卡片将通过 JavaScript 动态生成 -->
                            </div>
                        </div>
                    </div>
                </div>
            `;
            container.append(sectionHtml);
        });
    }

    // 根据送检单号查询
    function queryOrderByNumber(orderNumber, page = 1) {
        $.ajax({
            url: '/QmsSchedule/QueryOrderByNumber',
            type: 'POST',
            data: {
                orderNumber: orderNumber,
                page: page,
                pageSize: currentQueryPageSize
            },
            dataType: 'json',
            success: function(response) {
                if (response.status === 'success') {
                    showQueryResult(response.data);
                    showMachineOperationList(response.data);
                    currentQueryPage = page;
                } else {
                    showQueryError(response.msg);
                }
            },
            error: function(xhr, status, error) {
                showQueryError('查询失败: ' + error);
            }
        });
    }

    // 显示查询结果信息
    function showQueryResult(data) {
        let infoHtml;
        
        if (data.isCompleted) {
            // 已完成的情况 - 不显示"该机台共有X个排程中送检单"
            infoHtml = `
                <strong>查询结果：</strong>
                送检单号 <strong>${data.orderNumber}</strong> 当前在机台 <strong>${data.machineDesc}</strong>，
                送检单已完成
            `;
        } else {
            // 排程中或量测中的情况 - 保留"该机台共有X个排程中送检单"描述
            infoHtml = `
                <strong>查询结果：</strong>
                送检单号 <strong>${data.orderNumber}</strong> 当前在机台 <strong>${data.machineDesc}</strong>，
                排程中第 <strong>${data.currentRank}</strong> 位，该机台共有 <strong>${data.scheduledCount}</strong> 个排程中送检单
            `;
        }
        
        $('#queryInfoContent').html(infoHtml);
        $('#queryInfoAlert').removeClass('alert-danger').addClass('alert-success');
        $('#queryInfoSection').show();
    }

    // 显示查询错误信息
    function showQueryError(message) {
        const errorHtml = `
            <i class="fas fa-exclamation-triangle me-2"></i>
            <strong>查询失败：</strong>${message}
        `;
        $('#queryInfoContent').html(errorHtml);
        $('#queryInfoAlert').removeClass('alert-success').addClass('alert-danger');
        $('#queryInfoSection').show();
        $('#machineOperationSection').hide();
    }

    // 显示机台作业列表
    function showMachineOperationList(data) {
        // 更新机台作业标题
        $('#machineOperationTitle').html(`
            <i class="fas fa-cogs me-2"></i>机台作业列表 - ${data.machineDesc}
        `);
        $('#machineBadge').text(data.machineDesc);

        // 渲染表格数据
        renderMachineOperationTable(data.orders);
        
        // 更新分页信息
        updatePaginationInfo(data.pagination);
        
        // 显示机台作业区域
        $('#machineOperationSection').show();
    }

    // 渲染机台作业表格
    function renderMachineOperationTable(orders) {
        const tbody = $('#machineOperationTable tbody');
        tbody.empty();

        if (!orders || orders.length === 0) {
            tbody.append('<tr><td colspan="8" class="text-center text-muted">暂无数据</td></tr>');
            return;
        }

        orders.forEach(function(order, index) {
            // 检查是否是当前查询的送检单号
            const isCurrentQuery = order.QcRecordId.toString() === currentQueryOrderNumber.toString();
            const rowClass = isCurrentQuery ? 'table-warning' : '';
            const highlightIcon = isCurrentQuery ? '<i class="fas fa-search text-primary me-1"></i>' : '';
            
            const rowHtml = `
                <tr class="${rowClass}">
                    <td>${index + 1}</td>
                    <td>${highlightIcon}${order.QcRecordId}</td>
                    <td>${order.CreateByName || '-'}</td>
                    <td>${formatDateTime(order.ReceiveTime)}</td>
                    <td>${formatDateTime(order.StartTime)}</td>
                    <td>${formatDateTime(order.EndTime)}</td>
                    <td>
                        <span class="badge ${getStatusBadgeClass(order.Status)}">${order.StatusName}</span>
                    </td>
                    <td>${order.ExecutorName || '-'}</td>
                </tr>
            `;
            tbody.append(rowHtml);
        });
    }

    // 更新分页信息
    function updatePaginationInfo(pagination) {
        $('#totalRecords').text(pagination.totalCount);
        $('#currentPageQuery').text(`${pagination.currentPage}/${pagination.totalPages} 页`);
        
        // 更新分页按钮状态
        $('#prevPageQuery').prop('disabled', pagination.currentPage <= 1);
        $('#nextPageQuery').prop('disabled', pagination.currentPage >= pagination.totalPages);
    }

    // 根据机台类型加载机台状态
    function loadMachineStatusByType(machineTypeId) {
        // 隐藏所有机台类型区域
        hideAllMachineSections();
        
        // 显示选中的机台类型区域
        showMachineSection(machineTypeId);
        
        // 加载该类型的机台状态
        $.ajax({
            url: '/QmsSchedule/GetMachineStatusByType',
            type: 'POST',
            data: {
                qcMachineModeId: machineTypeId
            },
            dataType: 'json',
            success: function(response) {
                if (response.status === 'success') {
                    renderMachineStatusCards(machineTypeId, response.result);
                } else {
                    console.error('获取机台状态失败:', response.msg);
                }
            },
            error: function(xhr, status, error) {
                console.error('获取机台状态API调用失败:', error);
            }
        });
    }

    // 加载所有机台状态
    function loadAllMachineStatus() {
        // 显示所有机台类型区域
        showAllMachineSections();
        
        // 并行加载所有类型的机台状态
        allMachineTypes.forEach(function(type) {
            $.ajax({
                url: '/QmsSchedule/GetMachineStatusByType',
                type: 'POST',
                data: {
                    qcMachineModeId: type.QcMachineModeId
                },
                dataType: 'json',
                success: function(response) {
                    if (response.status === 'success') {
                        renderMachineStatusCards(type.QcMachineModeId, response.result);
                    } else {
                        console.error(`获取${type.QcMachineModeName}机台状态失败:`, response.msg);
                    }
                },
                error: function(xhr, status, error) {
                    console.error(`获取${type.QcMachineModeName}机台状态API调用失败:`, error);
                }
            });
        });
    }

    // 隐藏所有机台类型区域
    function hideAllMachineSections() {
        $('#machineStatusSections .col-12').hide();
    }

    // 显示所有机台类型区域
    function showAllMachineSections() {
        $('#machineStatusSections .col-12').show();
    }

    // 显示指定机台类型区域
    function showMachineSection(machineTypeId) {
        const sectionId = `#machineStatusSection${machineTypeId}`;
        $(sectionId).show();
    }

    // 渲染机台状态卡片
    function renderMachineStatusCards(machineTypeId, machines) {
        console.log(`渲染机台状态卡片 - 类型ID: ${machineTypeId}, 机台数量: ${machines ? machines.length : 0}`);
        
        const containerId = `#machineStatusCards${machineTypeId}`;
        const container = $(containerId);
        
        console.log(`容器ID: ${containerId}, 容器元素:`, container);
        
        container.empty();

        if (!machines || machines.length === 0) {
            console.log(`类型 ${machineTypeId} 暂无机台`);
            container.append('<div class="col-12 text-center text-muted">该类型暂无机台</div>');
            return;
        }

        console.log(`开始渲染 ${machines.length} 个机台卡片`);
        machines.forEach(function(machine) {
            const cardHtml = generateMachineStatusCard(machine);
            container.append(cardHtml);
        });
        console.log(`机台卡片渲染完成`);
    }

    // 生成机台状态卡片
    function generateMachineStatusCard(machine) {
        // 确定机台状态
        let measurementStatus = 'idle';
        let statusText = '空闲';
        let currentOrder = null;
        let currentOrderDisplay = '';

        if (machine.MeasuringCount > 0) {
            measurementStatus = 'measuring';
            statusText = '量测中';
            currentOrderDisplay = '<span class="text-primary">量测中</span>';
        } else if (machine.QueueCount > 0) {
            measurementStatus = 'queued';
            statusText = '排队中';
            currentOrderDisplay = '<span class="text-muted">等待中</span>';
        } else if (machine.CompletedCount > 0) {
            measurementStatus = 'completed';
            statusText = '已完成';
            currentOrderDisplay = '<span class="text-success">已完成</span>';
        } else {
            currentOrderDisplay = '<span class="text-muted">空闲</span>';
        }

        const cardHtml = `
            <div class="col-md-4 col-lg-2 mb-3">
                <div class="machine-card">
                    <div class="status-corner ${measurementStatus === 'measuring' || measurementStatus === 'queued' ? 'status-detecting' : 'status-standby'}">
                        ${statusText}
                    </div>
                    <div class="machine-header">
                        <div class="machine-title">${machine.MachineDesc}</div>
                    </div>
                    <div class="machine-status">
                        <div class="status-info">
                            <span class="status-label">排队数量</span>
                            <span class="status-value ${machine.QueueCount > 0 ? 'text-warning' : ''}">${machine.QueueCount}</span>
                        </div>
                        <div class="status-info">
                            <span class="status-label">当前检测</span>
                            <span class="status-value">${currentOrderDisplay}</span>
                        </div>
                    </div>
                </div>
            </div>
        `;
        return cardHtml;
    }

    // 获取状态徽章样式类
    function getStatusBadgeClass(status) {
        switch (status) {
            case '1': return 'bg-warning'; // 排程中
            case '2': return 'bg-info';    // 量测中
            case '3': return 'bg-success'; // 已完成
            default: return 'bg-secondary';
        }
    }

    // 格式化日期时间
    function formatDateTime(dateTimeStr) {
        if (!dateTimeStr) return '-';
        try {
            const date = new Date(dateTimeStr);
            return date.toLocaleString('zh-CN', {
                year: 'numeric',
                month: '2-digit',
                day: '2-digit',
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit'
            });
        } catch (e) {
            return dateTimeStr;
        }
    }

    // 显示提示信息
    function showAlert(message, type = 'info') {
        const alertClass = type === 'error' ? 'alert-danger' : 
                          type === 'success' ? 'alert-success' : 'alert-info';
        
        const alertHtml = `
            <div class="alert ${alertClass} alert-dismissible fade show" role="alert">
                ${message}
                <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
            </div>
        `;
        
        // 在页面顶部显示提示
        $('main').prepend(alertHtml);
        
        // 3秒后自动隐藏
        setTimeout(function() {
            $('.alert').fadeOut();
        }, 3000);
    }
    
    // 初始化用户信息显示
    function initializeUserDisplay() {
        try {
            const $display = $('#currentUserDisplay');
            const serverDisplay = (
                $display.attr('data-user-display') ||
                $display.data('userDisplay') ||
                ''
            ).toString().trim();

            if (serverDisplay.length > 0) {
                $display.text('当前登录：' + serverDisplay);
                return;
            }

            const userName = (
                $display.attr('data-user-name') ||
                $display.data('userName') ||
                ''
            ).toString().trim();
            const departmentName = (
                $display.attr('data-department-name') ||
                $display.data('departmentName') ||
                ''
            ).toString().trim();
            const companyName = (
                $display.attr('data-company-name') ||
                $display.data('companyName') ||
                ''
            ).toString().trim();

            if (userName.length > 0 || departmentName.length > 0 || companyName.length > 0) {
                const departmentPart = departmentName.length > 0 ? departmentName : '未分配部门';
                const companyPart = companyName.length > 0 ? companyName : '未分配公司';
                const composedDisplay = (userName.length > 0 ? userName : '未知用户') + '（' + departmentPart + ' / ' + companyPart + '）';
                $display.text('当前登录：' + composedDisplay);
                return;
            }

            // 从Cookie中读取Login字段
            const loginValue = getCookie('Login');
            if (loginValue) {
                $display.text('当前登录：' + loginValue);
            } else {
                $display.text('当前登录：未知用户');
            }
        } catch (error) {
            console.error('读取用户信息失败:', error);
            $('#currentUserDisplay').text('当前登录：加载失败');
        }
    }
    
    // 执行登出操作
    function performLogout() {
        // 防止重复调用
        if (window.isLoggingOut) {
            return;
        }
        window.isLoggingOut = true;
        
        if (confirm('确定要登出吗？')) {
            try {
                // 清除所有相关Cookie
                document.cookie = 'Login=; expires=Thu, 01 Jan 1970 00:00:00 UTC; path=/;';
                document.cookie = 'LoginKey=; expires=Thu, 01 Jan 1970 00:00:00 UTC; path=/;';
                document.cookie = 'ASP.NET_SessionId=; expires=Thu, 01 Jan 1970 00:00:00 UTC; path=/;';
                
                // 跳转到登录页面
                window.location.href = '/User/Login';
            } catch (error) {
                console.error('登出失败:', error);
                window.isLoggingOut = false; // 重置状态
            }
        } else {
            window.isLoggingOut = false; // 重置状态
        }
    }
    
    // 获取Cookie值的辅助函数
    function getCookie(name) {
        const value = `; ${document.cookie}`;
        const parts = value.split(`; ${name}=`);
        if (parts.length === 2) return parts.pop().split(';').shift();
        return null;
    }
});
