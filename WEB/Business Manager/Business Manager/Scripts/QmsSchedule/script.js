$(document).ready(function() {
    // 清理任何残留的模态框遮罩层
    function cleanupModalBackdrops() {
        $('.modal-backdrop').remove();
        $('body').removeClass('modal-open');
        $('body').css('padding-right', '');
    }
    
    // 页面加载时清理遮罩层
    cleanupModalBackdrops();
    
    // 监听模态框隐藏事件，确保遮罩层被清理
    $('.modal').on('hidden.bs.modal', function() {
        cleanupModalBackdrops();
    });

    // ==================== 全局登录弹窗功能 ====================
    
    // 显示登录弹窗
    function showLoginModal() {
        $('#loginModal').modal({
            backdrop: 'static',
            keyboard: false
        });
        $('#loginModal').modal('show');
        
        // 延迟聚焦到输入框
        setTimeout(function() {
            $('#loginEmployeeId').focus();
        }, 300);
    }

    // 验证员工工号是否存在于系统中
    function validateEmployeeId(employeeId) {
        // 直接调用真实API验证
        return validateEmployeeIdAPI(employeeId).then(function(apiEmployee) {
            if (apiEmployee) {
                // 将API返回的员工信息保存到本地缓存中
                employeeData[employeeId] = apiEmployee;
                return apiEmployee;
            }
            return null;
        }).catch(function(error) {
            console.error('员工验证API调用失败:', error);
            // 重新抛出错误，让上层处理
            throw error;
        });
    }
    
    // 真实API验证员工工号
    function validateEmployeeIdAPI(employeeId) {
        return new Promise(function(resolve, reject) {
            $.ajax({
                url: '/QmsSchedule/ValidateEmployee',
                type: 'POST',
                data: {
                    employeeNo: employeeId
                },
                dataType: 'json',
                success: function(response) {
                    if (response.status === 'success' && response.result) {
                        resolve(response.result);
                    } else {
                        // 当员工号不存在或其他错误时，抛出错误而不是返回null
                        reject(new Error(response.msg || '员工验证失败'));
                    }
                },
                error: function(xhr, status, error) {
                    reject(new Error('网络错误: ' + error));
                }
            });
        });
    }

    // 登录表单提交
    $('#loginForm').on('submit', function(e) {
        e.preventDefault();
        
        const employeeId = $('#loginEmployeeId').val().trim();
        const loginBtn = $('#loginBtn');
        const inputField = $('#loginEmployeeId');
        
        if (!employeeId) {
            inputField.addClass('is-invalid');
            return;
        }
        
        // 显示加载状态
        loginBtn.prop('disabled', true).html('<span class="spinner-border spinner-border-sm me-2"></span>验证中...');
        
        // 验证员工 - 现在总是返回Promise
        validateEmployeeId(employeeId).then(function(validEmployee) {
            handleLoginResult(validEmployee, employeeId, inputField, loginBtn);
        }).catch(function(error) {
            handleLoginError(error, inputField, loginBtn);
        });
    });
    
    // 处理登录结果
    function handleLoginResult(validEmployee, employeeId, inputField, loginBtn) {
            if (validEmployee) {
                // 验证成功
                inputField.removeClass('is-invalid').addClass('is-valid');
                loginBtn.removeClass('btn-primary').addClass('btn-success')
                        .html('<i class="fas fa-check-circle me-2"></i>登录成功');
                
                // 保存登录状态到sessionStorage
                sessionStorage.setItem('isLoggedIn', 'true');
                sessionStorage.setItem('loginEmployeeId', employeeId);
                sessionStorage.setItem('loginEmployeeName', validEmployee.UserName);
                sessionStorage.setItem('loginTime', new Date().toISOString());
                
                // 延迟关闭弹窗
                setTimeout(function() {
                    // 关闭登录弹窗
                    $('#loginModal').modal('hide');
                    
                    // 更新界面上的员工信息显示
                    updateEmployeeDisplay(employeeId, validEmployee.UserName);
                    
                    // 显示成功提示
                    showAlert(`欢迎使用量测系统！员工 ${validEmployee.UserName}(${employeeId}) 登录成功`, 'success');
                    
                    // 重置登录表单
                    resetLoginForm();
                }, 1000);
                
            } else {
                // 验证失败
                inputField.removeClass('is-valid').addClass('is-invalid');
            loginBtn.removeClass('btn-primary').addClass('btn-danger')
                    .html('<i class="fas fa-exclamation-circle me-2"></i>登录失败');
                
                setTimeout(function() {
                loginBtn.prop('disabled', false)
                        .removeClass('btn-danger').addClass('btn-primary')
                        .html('<i class="fas fa-sign-in-alt me-2"></i>员工登录');
                    inputField.focus().select();
                }, 1500);
            }
    }
    
    // 处理登录错误
    function handleLoginError(error, inputField, loginBtn) {
        inputField.removeClass('is-valid').addClass('is-invalid');
        
        // 显示具体的错误信息
        const errorMessage = error.message || error.msg || '验证失败';
        
        // 优化按钮错误状态显示
        loginBtn.removeClass('btn-primary').addClass('btn-danger')
                .html('<i class="fas fa-exclamation-circle me-2"></i>' + errorMessage);
        
        // 显示美观的错误提示
        showAlert(errorMessage, 'error');
        
        setTimeout(function() {
            loginBtn.prop('disabled', false)
                    .removeClass('btn-danger').addClass('btn-primary')
                    .html('<i class="fas fa-sign-in-alt me-2"></i>员工登录');
            inputField.focus().select();
        }, 2000);
    }

    // 重置登录表单
    function resetLoginForm() {
        $('#loginEmployeeId').val('').removeClass('is-valid is-invalid');
        $('#loginBtn').prop('disabled', false)
                .removeClass('btn-success btn-danger').addClass('btn-primary')
                .html('<i class="fas fa-sign-in-alt me-2"></i>员工登录');
    }

    // 登出功能
    $('#logoutBtn').click(function(e) {
        e.preventDefault();
        if (confirm('确定要登出吗？')) {
            // 清除登录状态
            sessionStorage.removeItem('isLoggedIn');
            sessionStorage.removeItem('loginEmployeeId');
            sessionStorage.removeItem('loginEmployeeName');
            sessionStorage.removeItem('loginTime');
            
            // 隐藏界面上的员工信息显示
            updateEmployeeDisplay('', '');
            
            showAlert('已成功登出', 'success');
            
            // 延迟显示登录弹窗
            setTimeout(function() {
                showLoginModal();
            }, 1000);
        }
    });

    // 检查登录状态
    function checkLoginStatus() {
        const isLoggedIn = sessionStorage.getItem('isLoggedIn');
        const loginEmployeeId = sessionStorage.getItem('loginEmployeeId');
        const loginEmployeeName = sessionStorage.getItem('loginEmployeeName');
        
        if (isLoggedIn === 'true' && loginEmployeeId && loginEmployeeName) {
            // 已登录，显示员工信息
            updateEmployeeDisplay(loginEmployeeId, loginEmployeeName);
            console.log('用户已登录，员工:', loginEmployeeName, '(', loginEmployeeId, ')');
        } else {
            // 未登录，显示登录弹窗
            updateEmployeeDisplay('', '');
            showLoginModal();
        }
    }

    // 更新界面上的员工信息显示
    function updateEmployeeDisplay(employeeId, employeeName) {
        if (employeeId && employeeName) {
            $('#currentEmployeeText').text(`员工: ${employeeName}(${employeeId})`);
            $('#currentEmployeeInfo').show();
        } else {
            $('#currentEmployeeInfo').hide();
        }
    }

    // 页面加载时检查登录状态
    checkLoginStatus();
    
    // 量测数据缓存
    const measurementData = {};

    // 当前机台数据
    let currentMachineData = [];
    let currentMachineId = '';
    let currentCompanyId = null;
    
    // 防重复加载标志
    let isLoadingData = false;
    
    // 从URL参数获取机台ID
    function getMachineIdFromUrl() {
        const urlParams = new URLSearchParams(window.location.search);
        const machineId = urlParams.get('id');
        return machineId ? parseInt(machineId) : null;
    }

    // 从URL参数或全局变量获取公司ID
    function getCompanyIdFromUrl() {
        if (typeof window.companyId !== 'undefined' && window.companyId !== null) {
            const parsed = parseInt(window.companyId, 10);
            if (!isNaN(parsed) && parsed > 0) {
                return parsed;
            }
        }

        const urlParams = new URLSearchParams(window.location.search);
        const companyIdParam = urlParams.get('companyId');
        const companyId = companyIdParam ? parseInt(companyIdParam, 10) : null;
        return !isNaN(companyId) && companyId > 0 ? companyId : null;
    }
    
    // 验证机器ID是否存在
    function validateMachineIdAPI(machineId) {
        return new Promise(function(resolve, reject) {
            $.ajax({
                url: '/QmsSchedule/GetMachineInfo',
                type: 'POST',
                data: {
                    machineId: machineId
                },
                dataType: 'json',
                success: function(response) {
                    if (response.status === 'success') {
                        resolve(response.result);
                    } else {
                        reject(new Error(response.msg || '机器验证失败'));
                    }
                },
                error: function(xhr, status, error) {
                    reject(new Error('机器验证API调用失败: ' + error));
                }
            });
        });
    }
    
    // 更新页面标题显示机台类型名称
    function updatePageTitle(machineInfo) {
        const machineTypeName = machineInfo.QcMachineModeName || machineInfo.QcMachineModeDesc || 'QV-' + machineInfo.QcMachineModeId;
        const titleText = `${machineTypeName} 量测机台状态`;
        
        // 更新标题
        $('h4.mb-3').text(titleText);
        
        console.log('页面标题已更新:', titleText);
    }
    
    // 初始化页面数据
    function initializePageData() {
        currentCompanyId = getCompanyIdFromUrl();
        if (!currentCompanyId) {
            console.warn('未指定公司ID，将使用默认公司(4)');
        }
        const machineId = getMachineIdFromUrl();
        if (machineId) {
            // 先验证机器ID是否存在
            validateMachineIdAPI(machineId)
                .then(function(machineInfo) {
                    currentMachineId = machineId;
                    // 更新页面显示的机器信息
                    $('#currentMachineId').text(machineInfo.MachineDesc || 'M' + machineId);
                    
                    // 更新标题显示机台类型名称和ID
                    updatePageTitle(machineInfo);
                    
                    console.log('机器验证成功:', machineInfo);
                    
                    // 调用真实API加载数据
                    loadMachineData(machineId);
                })
                .catch(function(error) {
                    console.error('机器验证失败:', error);
                    showAlert('机器ID验证失败: ' + error.message, 'error');
                    
                    // 清空表格数据，不显示任何测试数据
                    clearTableData();
                    clearMachineStatus();
                });
        } else {
            showAlert('未指定机台ID，请从正确链接访问', 'warning');
        }
    }
    
    // 加载机台数据
    function loadMachineData(machineId) {
        // 防止重复加载
        if (isLoadingData) {
            console.log('数据正在加载中，跳过重复请求');
            return;
        }
        
        isLoadingData = true;
        console.log('开始加载机台数据:', machineId);
        
        fetchMeasurementDataAPI(machineId)
            .then(function(response) {
                if (response.success) {
                    console.log('API返回数据:', response.data);
                    currentMachineData = response.data || [];
                    renderTableDataFromCache(machineId, 1);
                    updateMachineStatus();
                } else {
                    showAlert('加载机台数据失败: ' + response.message, 'error');
                }
            })
            .catch(function(error) {
                showAlert('加载机台数据失败: ' + error.message, 'error');
            })
            .finally(function() {
                isLoadingData = false;
                console.log('机台数据加载完成');
            });
    }
    
    // 分页相关变量
    let currentPage = 1;
    let pageSize = 5;
    let totalPages = 1;
    
    // 员工数据缓存
    const employeeData = {};

    // 机台状态数据缓存
    let machineStatusData = [];

    // ==================== API 接口 ====================
    
    // 获取量测数据API
    function fetchMeasurementDataAPI(machineId) {
        return new Promise(function(resolve, reject) {
            $.ajax({
                url: '/QmsSchedule/GetScheduleList',
                type: 'POST',
                data: {
                    machineId: machineId,
                    status: '',
                    pageIndex: 1,
                    pageSize: 100
                },
                dataType: 'json',
                success: function(response) {
                    if (response.status === 'success') {
                        resolve({
                            success: true,
                            data: response.data || [],
                            message: '获取数据成功'
                        });
                    } else {
                        reject({
                            success: false,
                            message: response.msg || '获取数据失败'
                        });
                    }
                },
                error: function(xhr, status, error) {
                    reject({
                        success: false,
                        message: '网络错误: ' + error
                    });
                }
            });
        });
    }

    // 添加订单到队列API
    function addOrderToQueueAPI(orderData) {
        return new Promise(function(resolve, reject) {
            // 调用真实的API
            $.ajax({
                url: '/QmsSchedule/AddToSchedule',
                type: 'POST',
                data: {
                    qcRecordId: orderData.orderNo,
                    machineId: currentMachineId,
                    userId: 1 // 这里应该从sessionStorage获取真实的userId
                },
                dataType: 'json',
                success: function(response) {
                    if (response.status === 'success') {
                    resolve({
                        success: true,
                            data: response.data,
                            message: response.msg || '订单添加成功'
                    });
                } else {
                    reject({
                        success: false,
                            message: response.msg || '添加失败',
                            error: 'API_ERROR'
                        });
                    }
                },
                error: function(xhr, status, error) {
                    reject({
                        success: false,
                        message: '网络错误: ' + error,
                        error: 'NETWORK_ERROR'
                    });
                }
            });
        });
    }


    // 移动机台API
    function moveMachineAPI(orderNo, fromMachine, toMachine) {
        return new Promise(function(resolve, reject) {
            // 调试信息：打印当前数据和查找条件
            console.log('移动机台调试信息:');
            console.log('查找订单号:', orderNo);
            console.log('当前机台数据:', currentMachineData);
            console.log('数据长度:', currentMachineData.length);
            
            // 找到对应的QmmId
            const order = currentMachineData.find(item => {
                const qcRecordMatch = item.QcRecordId && item.QcRecordId.toString() === orderNo;
                const orderNoMatch = item.orderNo && item.orderNo.toString() === orderNo;
                console.log('检查项目:', item, 'QcRecordId匹配:', qcRecordMatch, 'orderNo匹配:', orderNoMatch);
                return qcRecordMatch || orderNoMatch;
            });
            
            console.log('找到的订单:', order);
            
            if (!order) {
                reject(new Error('找不到对应的排程记录 - 订单不存在'));
                    return;
                }

            if (!order.QmmId) {
                reject(new Error('找不到对应的排程记录 - 缺少QmmId字段'));
                return;
            }
            
            const employeeId = sessionStorage.getItem('loginEmployeeId');
            const employeeInfo = employeeData[employeeId];
            const userId = employeeInfo ? employeeInfo.UserId : 1;
            
            $.ajax({
                url: '/QmsSchedule/ChangeMachine',
                type: 'POST',
                data: {
                    qmmId: order.QmmId,
                    toMachineId: toMachine,
                    userId: userId
                },
                dataType: 'json',
                success: function(response) {
                    if (response.status === 'success') {
                resolve({
                    success: true,
                            data: {
                                orderNo: orderNo,
                                fromMachine: fromMachine,
                                toMachine: toMachine
                            },
                            message: `订单 ${orderNo} 已成功移动到机台 ${toMachine}`
                        });
                    } else {
                        reject(new Error(response.msg || '移动机台失败'));
                    }
                },
                error: function(xhr, status, error) {
                    reject(new Error('网络错误: ' + error));
                }
            });
        });
    }

    // 开始量测API
    function startMeasurementAPI(qcRecordId) {
        // 找到对应的QmmId
        const order = currentMachineData.find(item => 
            (item.QcRecordId && item.QcRecordId.toString() === qcRecordId) || 
            (item.orderNo && item.orderNo.toString() === qcRecordId)
        );
        
        if (!order || !order.QmmId) {
            showAlert('找不到对应的排程记录', 'error');
            return;
        }
        
        const employeeId = sessionStorage.getItem('loginEmployeeId');
        const employeeInfo = employeeData[employeeId];
        const userId = employeeInfo ? employeeInfo.UserId : 1;
        
        $.ajax({
            url: '/QmsSchedule/StartMeasurement',
            type: 'POST',
            data: {
                qmmId: order.QmmId,
                executorId: userId
            },
            dataType: 'json',
            success: function(response) {
                if (response.status === 'success') {
                    showAlert('开始量测成功', 'success');
                    // 清空输入框
                    $('#orderNumber1').val('');
                    // 刷新数据
                    loadMachineData(currentMachineId);
                } else {
                    showAlert('开始量测失败: ' + response.msg, 'error');
                }
            },
            error: function(xhr, status, error) {
                showAlert('开始量测失败: ' + error, 'error');
            }
        });
    }

    // 结束量测API
    function endMeasurementAPI(qcRecordId) {
        // 找到对应的QmmId
        const order = currentMachineData.find(item => 
            (item.QcRecordId && item.QcRecordId.toString() === qcRecordId) || 
            (item.orderNo && item.orderNo.toString() === qcRecordId)
        );
        
        if (!order || !order.QmmId) {
            showAlert('找不到对应的排程记录', 'error');
            return;
        }
        
        const employeeId = sessionStorage.getItem('loginEmployeeId');
        const employeeInfo = employeeData[employeeId];
        const userId = employeeInfo ? employeeInfo.UserId : 1;
        
        $.ajax({
            url: '/QmsSchedule/EndMeasurement',
            type: 'POST',
                        data: {
                qmmId: order.QmmId,
                userId: userId
            },
            dataType: 'json',
            success: function(response) {
                if (response.status === 'success') {
                    showAlert('结束量测成功', 'success');
                    // 清空输入框
                    $('#orderNumber1').val('');
                    // 刷新数据
                    loadMachineData(currentMachineId);
                } else {
                    showAlert('结束量测失败: ' + response.msg, 'error');
                }
            },
            error: function(xhr, status, error) {
                showAlert('结束量测失败: ' + error, 'error');
            }
        });
    }

    // 获取当前机台的机型ID
    function getCurrentMachineModeId() {
        return new Promise(function(resolve, reject) {
            $.ajax({
                url: '/QmsSchedule/GetMachineInfo',
                type: 'POST',
                data: { machineId: currentMachineId },
                dataType: 'json',
                success: function(response) {
                    if (response.status === 'success' && response.result) {
                        resolve(response.result.QcMachineModeId);
                    } else {
                        reject(new Error(response.msg || '获取机台信息失败'));
                    }
                },
                error: function(xhr, status, error) {
                    reject(new Error('网络错误: ' + error));
                }
            });
        });
    }

    // 获取指定机型的所有机台API
    function getMachinesByModeAPI(qcMachineModeId) {
        return new Promise(function(resolve, reject) {
            $.ajax({
                url: '/QmsSchedule/GetMachinesByMode',
                type: 'POST',
                data: { qcMachineModeId: qcMachineModeId },
                dataType: 'json',
                success: function(response) {
                    if (response.status === 'success' && response.result) {
                        resolve(response.result);
                    } else {
                        reject(new Error(response.msg || '获取机台列表失败'));
                    }
                },
                error: function(xhr, status, error) {
                    reject(new Error('网络错误: ' + error));
                }
            });
        });
    }

    // 更新机台排队数量
    function updateMachineQueueCounts(machines) {
        machines.forEach(function(machine) {
            // 调用获取排程列表API来获取排队数量
            $.ajax({
                url: '/QmsSchedule/GetScheduleList',
                type: 'POST',
                data: { machineId: machine.MachineId },
                dataType: 'json',
                success: function(response) {
                    if (response.status === 'success' && response.data) {
                        const queueCount = response.data.length;
                        const queueElement = $(`.queue-number[data-machine-id="${machine.MachineId}"]`);
                        queueElement.text(queueCount);
                        
                        // 根据排队数量调整颜色
                        if (queueCount === 0) {
                            queueElement.css('color', '#28a745'); // 绿色
                        } else if (queueCount <= 3) {
                            queueElement.css('color', '#ffc107'); // 黄色
                        } else {
                            queueElement.css('color', '#dc3545'); // 红色
                        }
                    }
                },
                error: function() {
                    // 如果获取失败，显示问号
                    $(`.queue-number[data-machine-id="${machine.MachineId}"]`).text('?');
                }
            });
        });
    }

    // 获取机台状态API
    function fetchMachineStatusAPI(qcMachineModeId) {
        return new Promise(function(resolve, reject) {
            $.ajax({
                url: '/QmsSchedule/GetMachineStatus',
                type: 'POST',
                data: {
                    qcMachineModeId: qcMachineModeId || -1
                },
                dataType: 'json',
                success: function(response) {
                    if (response.status === 'success') {
                resolve({
                    success: true,
                            data: response.data || [],
                            message: '获取机台状态成功'
                        });
                    } else {
                        reject({
                            success: false,
                            message: response.msg || '获取机台状态失败'
                        });
                    }
                },
                error: function(xhr, status, error) {
                    reject({
                        success: false,
                        message: '网络错误: ' + error
                    });
                }
            });
        });
    }

    // ==================== 辅助函数 ====================
    
    // 检查订单是否在列表中
    function checkOrderInList(orderNumber) {
        const order = currentMachineData.find(item => 
            (item.QcRecordId && item.QcRecordId.toString() === orderNumber) || 
            (item.orderNo && item.orderNo.toString() === orderNumber)
        );
        return order !== undefined;
    }
    
    // 获取订单状态
    function getOrderStatus(orderNumber) {
        const order = currentMachineData.find(item => 
            (item.QcRecordId && item.QcRecordId.toString() === orderNumber) || 
            (item.orderNo && item.orderNo.toString() === orderNumber)
        );
        
        if (!order) return null;
        
        // 根据StartTime和EndTime动态计算状态
        if (order.EndTime && order.EndTime !== null) {
            return '量测完成';
        } else if (order.StartTime && order.StartTime !== null) {
            return '量测中';
        } else {
            return '排程中';
        }
    }

    // ==================== 数据渲染函数 ====================

    // 从缓存渲染表格数据（支持分页）
    function renderTableDataFromCache(machineId, page = 1) {
        const tbody = $('#orderTable tbody');
        tbody.empty();

        if (currentMachineData && currentMachineData.length > 0) {
            console.log('原始数据长度:', currentMachineData.length);
            console.log('原始数据:', currentMachineData);
            
            // 去重：按QcRecordId去重，保留最新的记录
            const seen = new Map();
            
            currentMachineData.forEach(function(item) {
                const key = item.QcRecordId;
                if (!seen.has(key)) {
                    // 第一次遇到这个QcRecordId，直接添加
                    seen.set(key, item);
                    console.log('添加新记录:', key, item);
                } else {
                    // 已经存在，比较时间决定是否替换
                    const existingItem = seen.get(key);
                    const currentTime = item.ReceiveTime || item.CreateDate || item.QcCreateDate;
                    const existingTime = existingItem.ReceiveTime || existingItem.CreateDate || existingItem.QcCreateDate;
                    
                    console.log('重复记录:', key, '当前时间:', currentTime, '已存在时间:', existingTime);
                    
                    if (currentTime && existingTime) {
                        // 如果都有时间，比较时间
                        if (new Date(currentTime) > new Date(existingTime)) {
                            seen.set(key, item);
                            console.log('替换为更新的记录:', key);
                        }
                    } else if (currentTime && !existingTime) {
                        // 如果当前项有时间而已存在项没有，替换
                        seen.set(key, item);
                        console.log('替换为有时间记录:', key);
                    }
                    // 如果当前项没有时间而已存在项有，保持已存在的项
                }
            });
            
            // 转换为数组
            const deduplicatedData = Array.from(seen.values());
            console.log('去重后数据长度:', deduplicatedData.length);
            console.log('去重后数据:', deduplicatedData);
            
            // 过滤掉已完成的订单（EndTime 不为空的数据）
            const filteredData = deduplicatedData.filter(function(item) {
                return !item.EndTime || item.EndTime === null;
            });
            console.log('过滤已完成订单后数据长度:', filteredData.length);
            console.log('过滤后数据:', filteredData);
            
            // 计算分页
            const totalItems = filteredData.length;
            totalPages = Math.ceil(totalItems / pageSize);
            currentPage = Math.min(page, totalPages);
            
            // 计算当前页的数据范围
            const startIndex = (currentPage - 1) * pageSize;
            const endIndex = startIndex + pageSize;
            const pageData = filteredData.slice(startIndex, endIndex);
            
            // 渲染当前页数据
            pageData.forEach(function(order, index) {
                // 映射后端字段到前端显示
                const orderNo = order.QcRecordId || order.orderNo;
                const inspector = order.CreateByName || order.inspector || '-';
                const submitTime = order.ReceiveTime ? new Date(order.ReceiveTime).toLocaleString() : (order.submitTime || '-');
                const startTime = order.StartTime ? new Date(order.StartTime).toLocaleString() : (order.startTime || '-');
                const endTime = order.EndTime ? new Date(order.EndTime).toLocaleString() : (order.endTime || '-');
                
                // 根据StartTime和EndTime动态计算状态
                let status = '未知';
                let statusClass = '';
                let statusBadge = '';
                
                if (order.EndTime && order.EndTime !== null) {
                    // 3.量测完成(EndTime is not null)
                    status = '量测完成';
                    statusClass = 'table-success';
                    statusBadge = '<span class="badge bg-success">量测完成</span>';
                } else if (order.StartTime && order.StartTime !== null) {
                    // 2.量測中(StartTime is not null && EndTime is null)
                    status = '量测中';
                    statusClass = 'table-info';
                    statusBadge = '<span class="badge bg-primary">量测中</span>';
                } else {
                    // 1.排程中(StartTime is null)
                    status = '排程中';
                    statusClass = 'table-warning';
                    statusBadge = '<span class="badge bg-warning">排程中</span>';
                }
                
                const row = `
                    <tr class="${statusClass}">
                        <td class="text-center align-middle">
                            <input class="form-check-input row-checkbox" type="checkbox" value="${orderNo}">
                        </td>
                        <td>${orderNo}</td>
                        <td>${inspector}</td>
                        <td>${submitTime}</td>
                        <td>${startTime}</td>
                        <td>${endTime}</td>
                        <td>${statusBadge}</td>
                        <td>
                            <button class="btn btn-sm btn-outline-primary move-machine-btn" data-order-no="${orderNo}">移动机台</button>
                        </td>
                    </tr>
                `;
                tbody.append(row);
            });
            
            // 更新分页信息
            updatePaginationInfo(totalItems, currentPage);
        } else {
            tbody.append('<tr><td colspan="8" class="text-center text-muted">暂无数据</td></tr>');
            updatePaginationInfo(0, 1);
        }
        
        // 更新移动机台按钮状态
        updateMoveMachineButtonState();
    }

    // 渲染表格数据（支持分页）- 重新获取数据
    function renderTableData(machineId, page = 1) {
        // 防止重复调用
        if (isLoadingData) {
            console.log('数据正在加载中，跳过重复渲染请求');
            return;
        }
        
        const tbody = $('#orderTable tbody');
        tbody.empty();
        
        console.log('开始重新获取数据并渲染:', machineId, '页码:', page);

        fetchMeasurementDataAPI(machineId)
            .then(function(response) {
                if (response.success && response.data.length > 0) {
                    // 更新缓存数据
                    currentMachineData = response.data || [];
                    
                    // 去重：按QcRecordId去重，保留最新的记录
                    const seen = new Map();
                    
                    response.data.forEach(function(item) {
                        const key = item.QcRecordId;
                        if (!seen.has(key)) {
                            // 第一次遇到这个QcRecordId，直接添加
                            seen.set(key, item);
                        } else {
                            // 已经存在，比较时间决定是否替换
                            const existingItem = seen.get(key);
                            const currentTime = item.ReceiveTime || item.CreateDate || item.QcCreateDate;
                            const existingTime = existingItem.ReceiveTime || existingItem.CreateDate || existingItem.QcCreateDate;
                            
                            if (currentTime && existingTime) {
                                // 如果都有时间，比较时间
                                if (new Date(currentTime) > new Date(existingTime)) {
                                    seen.set(key, item);
                                }
                            } else if (currentTime && !existingTime) {
                                // 如果当前项有时间而已存在项没有，替换
                                seen.set(key, item);
                            }
                            // 如果当前项没有时间而已存在项有，保持已存在的项
                        }
                    });
                    
                    // 转换为数组
                    const deduplicatedData = Array.from(seen.values());
                    
                    // 过滤掉已完成的订单（EndTime 不为空的数据）
                    const filteredData = deduplicatedData.filter(function(item) {
                        return !item.EndTime || item.EndTime === null;
                    });
                    
                    // 计算分页
                    const totalItems = filteredData.length;
                    totalPages = Math.ceil(totalItems / pageSize);
                    currentPage = Math.min(page, totalPages);
                    
                    // 计算当前页的数据范围
                    const startIndex = (currentPage - 1) * pageSize;
                    const endIndex = startIndex + pageSize;
                    const pageData = filteredData.slice(startIndex, endIndex);
                    
                    // 渲染当前页数据
                    pageData.forEach(function(order, index) {
                        // 映射后端字段到前端显示
                        const orderNo = order.QcRecordId || order.orderNo;
                        const inspector = order.CreateByName || order.inspector || '-';
                        const submitTime = order.ReceiveTime ? new Date(order.ReceiveTime).toLocaleString() : (order.submitTime || '-');
                        const startTime = order.StartTime ? new Date(order.StartTime).toLocaleString() : (order.startTime || '-');
                        const endTime = order.EndTime ? new Date(order.EndTime).toLocaleString() : (order.endTime || '-');
                        
                        // 根据StartTime和EndTime动态计算状态
                        let status = '未知';
                        let statusClass = '';
                        let statusBadge = '';
                        
                        if (order.EndTime && order.EndTime !== null) {
                            // 3.量测完成(EndTime is not null)
                            status = '量测完成';
                            statusClass = 'table-success';
                            statusBadge = '<span class="badge bg-success">量测完成</span>';
                        } else if (order.StartTime && order.StartTime !== null) {
                            // 2.量测中(StartTime is not null && EndTime is null)
                            status = '量测中';
                            statusClass = 'table-info';
                            statusBadge = '<span class="badge bg-primary">量测中</span>';
                        } else {
                            // 1.排程中(StartTime is null)
                            status = '排程中';
                            statusClass = 'table-warning';
                            statusBadge = '<span class="badge bg-warning">排程中</span>';
                        }

                        // 急件样式（暂时没有急件字段）
                        if (order.urgent) {
                            statusClass += ' table-danger';
                        }

                        const row = `
                            <tr class="${statusClass}">
                                <td class="text-center align-middle">
                                    <input class="form-check-input row-checkbox" type="checkbox" value="${orderNo}">
                                </td>
                                <td>
                                    ${orderNo}
                                    ${order.urgent ? '<span class="badge bg-danger ms-1">急件</span>' : ''}
                                </td>
                                <td>${inspector}</td>
                                <td>${submitTime}</td>
                                <td>${startTime}</td>
                                <td>${endTime}</td>
                                <td>${statusBadge}</td>
                                <td>
                                    <button class="btn btn-sm btn-outline-primary move-machine-btn" data-order-no="${orderNo}">移动机台</button>
                                </td>
                            </tr>
                        `;
                        tbody.append(row);
                    });
                    
                    // 更新分页信息
                    updatePaginationInfo(totalItems, currentPage);
                } else {
                    tbody.append('<tr><td colspan="8" class="text-center text-muted">暂无数据</td></tr>');
                    updatePaginationInfo(0, 1);
                }
                
                // 更新移动机台按钮状态
                updateMoveMachineButtonState();
            })
            .catch(function(error) {
                console.error('获取量测数据失败:', error);
                tbody.append('<tr><td colspan="8" class="text-center text-danger">加载数据失败</td></tr>');
                updatePaginationInfo(0, 1);
                
                // 更新移动机台按钮状态
                updateMoveMachineButtonState();
            });
    }

    // 更新分页信息
    function updatePaginationInfo(totalItems, currentPage) {
        $('#totalRecords').text(totalItems);
        generatePagination(currentPage, totalPages);
    }

    // 生成Bootstrap分页组件
    function generatePagination(currentPage, totalPages) {
        const paginationNav = $('#paginationNav');
        paginationNav.empty();
        
        console.log(`生成分页: 当前页=${currentPage}, 总页数=${totalPages}`);
        
        if (totalPages <= 1) {
            return; // 只有一页时不显示分页
        }
        
        // 计算显示的页码范围
        let startPage = Math.max(1, currentPage - 2);
        let endPage = Math.min(totalPages, currentPage + 2);
        
        // 确保显示5个页码（如果总页数足够）
        if (endPage - startPage < 4) {
            if (startPage === 1) {
                endPage = Math.min(totalPages, startPage + 4);
            } else {
                startPage = Math.max(1, endPage - 4);
            }
        }
        
        // 上一页按钮
        const prevItem = $(`
            <li class="page-item ${currentPage <= 1 ? 'disabled' : ''}">
                <a class="page-link" href="#" data-page="${currentPage - 1}">
                    <i class="fas fa-chevron-left"></i>
                </a>
            </li>
        `);
        paginationNav.append(prevItem);
        
        // 第一页和省略号
        if (startPage > 1) {
            const firstItem = $(`
                <li class="page-item">
                    <a class="page-link" href="#" data-page="1">1</a>
                </li>
            `);
            paginationNav.append(firstItem);
            
            if (startPage > 2) {
                const ellipsisItem = $(`
                    <li class="page-item disabled">
                        <span class="page-link">...</span>
                    </li>
                `);
                paginationNav.append(ellipsisItem);
            }
        }
        
        // 页码按钮
        for (let i = startPage; i <= endPage; i++) {
            const pageItem = $(`
                <li class="page-item ${i === currentPage ? 'active' : ''}">
                    <a class="page-link" href="#" data-page="${i}">${i}</a>
                </li>
            `);
            paginationNav.append(pageItem);
        }
        
        // 最后一页和省略号
        if (endPage < totalPages) {
            if (endPage < totalPages - 1) {
                const ellipsisItem = $(`
                    <li class="page-item disabled">
                        <span class="page-link">...</span>
                    </li>
                `);
                paginationNav.append(ellipsisItem);
            }
            
            const lastItem = $(`
                <li class="page-item">
                    <a class="page-link" href="#" data-page="${totalPages}">${totalPages}</a>
                </li>
            `);
            paginationNav.append(lastItem);
        }
        
        // 下一页按钮
        const nextItem = $(`
            <li class="page-item ${currentPage >= totalPages ? 'disabled' : ''}">
                <a class="page-link" href="#" data-page="${currentPage + 1}">
                    <i class="fas fa-chevron-right"></i>
                </a>
            </li>
        `);
        paginationNav.append(nextItem);
    }

    // 生成机台状态卡片
    function generateMachineStatusCards() {
        const container = $('#machineStatusCards');
        container.empty();
        
        // 显示加载状态
        container.html('<div class="col-12 text-center text-muted"><i class="fas fa-spinner fa-spin me-2"></i>加载中...</div>');

        // 获取当前机台的机型ID，然后获取同类型的所有机台
        getCurrentMachineModeId()
            .then(function(modeId) {
                if (!modeId) {
                    console.warn('无法获取当前机台类型');
                    container.empty(); // 清空加载提示
                    return;
                }
                
                // 获取同类型的所有机台
                return getMachinesByModeAPI(modeId);
            })
            .then(function(machines) {
                container.empty(); // 清空加载提示
                
                if (!machines || machines.length === 0) {
                    container.html('<div class="col-12 text-center text-muted">没有找到机台</div>');
                    return;
                }
                
                // 为每个机台生成状态卡片
                machines.forEach(function(machine) {
                    generateMachineStatusCard(machine);
                });
            })
            .catch(function(error) {
                console.error('获取机台列表失败:', error);
                container.empty(); // 清空加载提示，静默处理错误
            });
    }

    // 生成单个机台状态卡片
    function generateMachineStatusCard(machine) {
        const container = $('#machineStatusCards');
        
        // 获取机台的量测数据状态
        getMachineMeasurementStatus(machine.MachineId)
            .then(function(status) {
            const cardHtml = `
                <div class="col-md-4 col-lg-2 mb-3">
                        <div class="machine-card ${machine.MachineId.toString() === currentMachineId ? 'current-machine' : ''}">
                            <div class="status-corner ${status.measurementStatus === 'measuring' || status.measurementStatus === 'queued' ? 'status-detecting' : 'status-standby'}">
                                ${status.statusText}
                        </div>
                        <div class="machine-header">
                                <div class="machine-title">${machine.MachineDesc}</div>
                        </div>
                        <div class="machine-status">
                            <div class="status-info">
                                <span class="status-label">排队数量</span>
                                    <span class="status-value">${status.queueCount}</span>
                            </div>
                            <div class="status-info">
                                <span class="status-label">当前检测</span>
                                    <span class="status-value">${status.currentOrder || '-'}</span>
                                </div>
                            </div>
                        </div>
                    </div>
                `;
                container.append(cardHtml);
            })
            .catch(function(error) {
                console.error('获取机台状态失败:', error);
                // 即使获取状态失败，也显示基本的机台卡片
                const cardHtml = `
                    <div class="col-md-4 col-lg-2 mb-3">
                        <div class="machine-card ${machine.MachineId.toString() === currentMachineId ? 'current-machine' : ''}">
                            <div class="status-corner status-standby">空闲</div>
                            <div class="machine-header">
                                <div class="machine-title">${machine.MachineDesc}</div>
                            </div>
                            <div class="machine-status">
                                <div class="status-info">
                                    <span class="status-label">排队数量</span>
                                    <span class="status-value">-</span>
                                </div>
                                <div class="status-info">
                                    <span class="status-label">当前检测</span>
                                    <span class="status-value">-</span>
                            </div>
                        </div>
                    </div>
                </div>
            `;
            container.append(cardHtml);
        });
    }

    // 获取机台量测数据状态
    function getMachineMeasurementStatus(machineId) {
        return new Promise(function(resolve) {
            // 调用API获取机台的量测数据
            fetchMeasurementDataAPI(machineId)
                .then(function(response) {
                    if (response.success && response.data && response.data.length > 0) {
                        const data = response.data;
                        
                        // 统计排队数量（排程中的订单）
                        const queueCount = data.filter(item => 
                            !item.StartTime || item.StartTime === null
                        ).length;
                        
                        // 统计正在量测的订单数量（有开始时间但没有结束时间）
                        const measuringCount = data.filter(item => 
                            item.StartTime && item.StartTime !== null && 
                            (!item.EndTime || item.EndTime === null)
                        ).length;
                        
                        // 检查是否有已完成的订单
                        const completedOrder = data.find(item => 
                            item.EndTime && item.EndTime !== null
                        );
                        
                        let measurementStatus = 'idle';
                        let statusText = '空闲';
                        let currentOrder = null;
                        
                        if (measuringCount > 0) {
                            measurementStatus = 'measuring';
                            statusText = '量测中';
                            currentOrder = measuringCount; // 显示正在检测的数量
                        } else if (queueCount > 0) {
                            measurementStatus = 'queued';
                            statusText = '排队中';
                        } else if (completedOrder) {
                            measurementStatus = 'completed';
                            statusText = '已完成';
                        }
                        
                        resolve({
                            isOnline: true,
                            queueCount: queueCount,
                            measurementStatus: measurementStatus,
                            statusText: statusText,
                            currentOrder: currentOrder
                        });
                    } else {
                        // 没有数据，机台空闲
                        resolve({
                            isOnline: true,
                            queueCount: 0,
                            measurementStatus: 'idle',
                            statusText: '空闲',
                            currentOrder: null
                        });
                    }
                })
                .catch(function(error) {
                    console.error('获取机台量测数据失败:', error);
                    resolve({
                        isOnline: false,
                        queueCount: 0,
                        measurementStatus: 'unknown',
                        statusText: '离线',
                        currentOrder: null
                    });
                });
        });
    }

    // 初始化机台状态卡片
    generateMachineStatusCards();
    
    
    // 初始化移动机台按钮状态
    updateMoveMachineButtonState();

    // 更新移动机台按钮状态
    function updateMoveMachineButtonState() {
        const checkedRows = $('.row-checkbox:checked').length;
        const moveBtn = $('#moveMachineBtn');
        
        if (checkedRows > 0) {
            moveBtn.prop('disabled', false);
            moveBtn.removeClass('btn-outline-secondary').addClass('btn-outline-primary');
        } else {
            moveBtn.prop('disabled', true);
            moveBtn.removeClass('btn-outline-primary').addClass('btn-outline-secondary');
        }
    }

    // 全选/取消全选功能
    $('#selectAll').change(function() {
        const isChecked = $(this).is(':checked');
        $('.row-checkbox').prop('checked', isChecked);
        updateMoveMachineButtonState();
    });

    // 行复选框变化时更新全选状态
    $(document).on('change', '.row-checkbox', function() {
        const totalCheckboxes = $('.row-checkbox').length;
        const checkedCheckboxes = $('.row-checkbox:checked').length;
        
        $('#selectAll').prop('checked', totalCheckboxes === checkedCheckboxes);
        $('#selectAll').prop('indeterminate', checkedCheckboxes > 0 && checkedCheckboxes < totalCheckboxes);
        
        // 更新移动机台按钮状态
        updateMoveMachineButtonState();
    });

    // 添加到队列按钮 - 显示输入送检单号弹窗（独立功能）
    $('#addToQueueBtn').click(function(e) {
        e.preventDefault();
        e.stopPropagation();
        
        console.log('添加到队列按钮被点击');
        
        // 清空之前的输入
        $('#modalOrderNumberInput').val('').removeClass('is-invalid is-valid');
        
        // 清理可能存在的遮罩层
        cleanupModalBackdrops();
        
        // 显示输入弹窗
        $('#inputOrderNumberModal').modal({
            backdrop: 'static',
            keyboard: true,
            show: true
        });
        
        // 延迟聚焦到输入框
        setTimeout(function() {
            $('#modalOrderNumberInput').focus();
        }, 300);
    });

    // 输入弹窗 - 查询按钮点击事件
    $('#modalQueryBtn').click(function() {
        handleModalOrderQuery();
    });

    // 输入弹窗 - 表单提交事件（支持回车键）
    $('#inputOrderNumberForm').on('submit', function(e) {
        e.preventDefault();
        handleModalOrderQuery();
    });

    // 输入弹窗 - 输入框回车事件
    $('#modalOrderNumberInput').keypress(function(e) {
        if (e.which === 13) {
            e.preventDefault();
            handleModalOrderQuery();
        }
    });

    // 处理弹窗中的订单查询
    function handleModalOrderQuery() {
        const orderNumber = $('#modalOrderNumberInput').val().trim();
        const inputField = $('#modalOrderNumberInput');
        const queryBtn = $('#modalQueryBtn');
        
        if (!orderNumber) {
            inputField.addClass('is-invalid');
            return;
        }
        
        // 移除无效状态
        inputField.removeClass('is-invalid');
        
        // 显示加载状态
        queryBtn.prop('disabled', true).html('<span class="spinner-border spinner-border-sm me-2"></span>查询中...');
        
        // 调用查询API
        searchQcRecordByIdFromModal(orderNumber, inputField, queryBtn);
    }

    // 从弹窗查询量测单号（查询成功后关闭输入弹窗）
    function searchQcRecordByIdFromModal(qcRecordId, inputField, queryBtn) {
        searchQcRecordByIdAPI(qcRecordId, currentCompanyId)
            .then(function(qcRecord) {
                console.log('查询到量测单据:', qcRecord);
                
                // 检查量测状态
                if (qcRecord.CheckQcMeasureData === 'F') {
                    inputField.addClass('is-invalid');
                    showAlert('该量测单已被驳回，无法添加到队列', 'warning');
                    queryBtn.prop('disabled', false).html('<i class="fas fa-search me-1"></i>查询');
                    return;
                }
                
                if (qcRecord.CheckQcMeasureData === 'E') {
                    inputField.addClass('is-invalid');
                    showAlert('该量测单已完成，无法添加到队列', 'warning');
                    queryBtn.prop('disabled', false).html('<i class="fas fa-search me-1"></i>查询');
                    return;
                }
                
                // 查询成功 - 显示成功状态
                inputField.removeClass('is-invalid').addClass('is-valid');
                queryBtn.removeClass('btn-primary').addClass('btn-success')
                        .html('<i class="fas fa-check-circle me-2"></i>查询成功');
                
                // 延迟关闭输入弹窗并显示处理弹窗
                setTimeout(function() {
                    // 关闭输入弹窗
                    $('#inputOrderNumberModal').modal('hide');
                    
                    // 重置查询按钮状态
                    queryBtn.prop('disabled', false)
                            .removeClass('btn-success').addClass('btn-primary')
                            .html('<i class="fas fa-search me-1"></i>查询');
                    
                    // 显示接收/驳回弹窗，并传入完整的量测单据信息
                    showReceiveRejectModal(qcRecord);
                }, 800);
            })
            .catch(function(error) {
                console.error('查询量测单号失败:', error);
                inputField.addClass('is-invalid');
                showAlert('查询失败: ' + error.message, 'error');
                
                // 重置查询按钮状态
                queryBtn.prop('disabled', false).html('<i class="fas fa-search me-1"></i>查询');
            });
    }

    // 开始/结束量测按钮
    $('#toggleMeasurementBtn').click(function() {
        const orderNumber = $('#orderNumber1').val().trim();
        
        if (!orderNumber) {
            showAlert('请输入送检单号', 'warning');
            return;
        }
        
        // 检查送检单号是否在列表中
        const orderExists = checkOrderInList(orderNumber);
        
        if (!orderExists) {
            showAlert('请先加入到列表', 'warning');
            return;
        }
        
        // 获取订单状态
        const orderStatus = getOrderStatus(orderNumber);
        
        if (orderStatus === '量测完成') {
            showAlert('该订单已完成，无法操作', 'warning');
            return;
        }
        
        // 根据当前状态决定是开始还是结束量测
        if (orderStatus === '排程中') {
            // 开始量测
            startMeasurementAPI(orderNumber);
        } else if (orderStatus === '量测中') {
            // 结束量测
            endMeasurementAPI(orderNumber);
        } else {
            showAlert('未知状态，无法操作', 'error');
        }
    });

    // 移动机台按钮（表格中的）
    $(document).on('click', '.move-machine-btn', function() {
        const orderNumber = $(this).data('order-no');
        showMoveMachineModal(orderNumber);
    });

    // 移动机台按钮（单独的）
    $('#moveMachineBtn').click(function() {
        const selectedOrders = getSelectedOrders();
        if (selectedOrders.length === 0) {
            showAlert('请选择要移动的订单', 'warning');
            return;
        }
        
        // 如果只选择了一个订单，显示单个移动机台模态框
        if (selectedOrders.length === 1) {
            showMoveMachineModal(selectedOrders[0]);
        } else {
            // 多个订单的情况，显示批量移动模态框
            showBatchMoveMachineModal(selectedOrders);
        }
    });

    // 机台选择卡片点击事件
    $(document).on('click', '.machine-selection-card:not(.disabled)', function() {
        $('.machine-selection-card').removeClass('selected');
        $(this).addClass('selected');
        $('#confirmMoveMachine').prop('disabled', false);
    });

    // 批量移动机台选择卡片点击事件
    $(document).on('click', '#batchMachineSelectionCards .machine-selection-card:not(.disabled)', function() {
        $('#batchMachineSelectionCards .machine-selection-card').removeClass('selected');
        $(this).addClass('selected');
        $('#confirmBatchMoveMachine').prop('disabled', false);
    });

    // 确认移动机台按钮
    $('#confirmMoveMachine').click(function() {
        const selectedMachine = $('.machine-selection-card.selected');
        if (selectedMachine.length === 0) {
            showAlert('请选择目标机台', 'warning');
            return;
        }
        
        const targetMachine = selectedMachine.data('machine-id');
        const orderNumber = $('#moveMachineModal').data('order-number');
        
        moveMachineToTarget(orderNumber, targetMachine);
    });

    // 确认批量移动机台按钮
    $('#confirmBatchMoveMachine').click(function() {
        const selectedMachine = $('#batchMachineSelectionCards .machine-selection-card.selected');
        if (selectedMachine.length === 0) {
            showAlert('请选择目标机台', 'warning');
            return;
        }
        
        const targetMachine = selectedMachine.data('machine-id');
        const selectedOrders = $('#batchMoveMachineModal').data('selected-orders');
        
        batchMoveMachinesToTarget(selectedOrders, targetMachine);
    });

    // 取消按钮点击事件
    $('button[data-dismiss="modal"]').click(function() {
        const modal = $(this).closest('.modal');
        modal.modal('hide');
    });

    // 分页功能 - Bootstrap分页组件
    $(document).on('click', '.page-link', function(e) {
        e.preventDefault();
        
        const targetPage = parseInt($(this).data('page'));
        console.log(`分页点击: 目标页=${targetPage}, 当前页=${currentPage}, 总页数=${totalPages}`);
        
        if (targetPage && targetPage !== currentPage && targetPage >= 1 && targetPage <= totalPages) {
            renderTableDataFromCache(currentMachineId, targetPage);
        }
    });

    // 刷新按钮
    $('#refreshBtn').click(function() {
        refreshData();
    });


    // 辅助函数
    function addToQueue(orderNumber) {
        const orderData = {
            orderNo: orderNumber,
            inspector: '当前用户',
            remark: '',
            urgent: false
        };
        
        addOrderToQueueAPI(orderData)
            .then(function(response) {
                showAlert(response.message, 'success');
                // 刷新表格数据
                renderTableData(currentMachineId, currentPage);
                // 更新机台状态
                updateMachineStatus();
            })
            .catch(function(error) {
                showAlert(error.message, 'error');
            });
    }


    // 显示移动机台模态框
    function showMoveMachineModal(orderNumber) {
        // 设置订单号到模态框
        $('#moveMachineModal').data('order-number', orderNumber);
        
        // 更新模态框信息
        $('#moveMachineInfo').html(`
            送检单号: <strong>${orderNumber}</strong> 当前所在机台: <strong>${currentMachineId}</strong>, 
            请选择将移动 1 笔送检单到目标机台
        `);
        
        // 生成机台选择卡片
        generateMachineSelectionCards();
        
        // 重置选择状态
        $('#confirmMoveMachine').prop('disabled', true);
        $('.machine-selection-card').removeClass('selected');
        
        // 显示模态框
        $('#moveMachineModal').modal('show');
    }

    // 显示批量移动机台模态框
    function showBatchMoveMachineModal(selectedOrders) {
        // 设置选中的订单到模态框
        $('#batchMoveMachineModal').data('selected-orders', selectedOrders);
        
        // 更新模态框信息
        const orderList = selectedOrders.join(', ');
        $('#batchMoveMachineInfo').html(`
            已选择 <strong>${selectedOrders.length}</strong> 笔送检单: <strong>${orderList}</strong><br>
            当前所在机台: <strong>${currentMachineId}</strong>, 
            请选择将移动这些送检单到目标机台
        `);
        
        // 生成机台选择卡片
        generateBatchMachineSelectionCards();
        
        // 重置选择状态
        $('#confirmBatchMoveMachine').prop('disabled', true);
        $('.machine-selection-card').removeClass('selected');
        
        // 显示模态框
        $('#batchMoveMachineModal').modal('show');
    }

    // 生成机台选择卡片
    function generateMachineSelectionCards() {
        const container = $('#machineSelectionCards');
        container.empty();

        // 显示加载状态
        container.html('<div class="col-12 text-center"><div class="spinner-border" role="status"><span class="sr-only">加载中...</span></div></div>');

        // 获取当前机台的机型ID
        getCurrentMachineModeId()
            .then(function(modeId) {
                if (!modeId) {
                    console.warn('无法获取当前机台类型');
                    container.empty(); // 清空加载提示，静默处理
                    return;
                }
                
                // 获取同类型的所有机台
                return getMachinesByModeAPI(modeId);
            })
            .then(function(machines) {
                if (!machines || machines.length === 0) {
                    container.html('<div class="col-12 text-center text-muted">没有找到同类型的机台</div>');
                    return;
                }
                
                // 过滤掉当前机台
                const availableMachines = machines.filter(machine => 
                    machine.MachineId.toString() !== currentMachineId.toString()
                );
                
                if (availableMachines.length === 0) {
                    container.html('<div class="col-12 text-center text-muted">没有其他同类型机台可移动</div>');
                    return;
                }
                
                // 清空加载状态，生成机台选择卡片
                container.empty();
        availableMachines.forEach(function(machine) {
                    generateMachineSelectionCardForMove(machine, container);
                });
            })
            .catch(function(error) {
                console.error('获取机台列表失败:', error);
                // 不显示错误提示，静默处理
            });
    }

    // 生成移动机台选择卡片（与状态卡片使用相同逻辑）
    function generateMachineSelectionCardForMove(machine, container) {
        // 获取机台的量测数据状态
        getMachineMeasurementStatus(machine.MachineId)
            .then(function(status) {
            const cardHtml = `
                    <div class="machine-selection-card" data-machine-id="${machine.MachineId}">
                        <div class="status-corner ${status.measurementStatus === 'measuring' || status.measurementStatus === 'queued' ? 'status-detecting' : 'status-standby'}">
                            ${status.statusText}
                        </div>
                        <div class="machine-header">
                            <div class="machine-title">${machine.MachineDesc}</div>
                        </div>
                        <div class="machine-status">
                            <div class="status-info">
                                <span class="status-label">排队数量</span>
                                <span class="status-value">${status.queueCount}</span>
                            </div>
                            <div class="status-info">
                                <span class="status-label">当前检测</span>
                                <span class="status-value">${status.currentOrder || '-'}</span>
                            </div>
                        </div>
                    </div>
                `;
                container.append(cardHtml);
            })
            .catch(function(error) {
                console.error('获取机台状态失败:', error);
                // 即使获取状态失败，也显示基本的机台卡片
                const cardHtml = `
                    <div class="machine-selection-card" data-machine-id="${machine.MachineId}">
                        <div class="status-corner status-standby">空闲</div>
                        <div class="machine-header">
                            <div class="machine-title">${machine.MachineDesc}</div>
                        </div>
                        <div class="machine-status">
                            <div class="status-info">
                                <span class="status-label">排队数量</span>
                                <span class="status-value">-</span>
                            </div>
                            <div class="status-info">
                                <span class="status-label">当前检测</span>
                                <span class="status-value">-</span>
                            </div>
                    </div>
                </div>
            `;
            container.append(cardHtml);
        });
    }

    // 生成批量移动机台选择卡片
    function generateBatchMachineSelectionCards() {
        const container = $('#batchMachineSelectionCards');
        container.empty();
        
        // 显示加载状态
        container.html('<div class="col-12 text-center"><div class="spinner-border" role="status"><span class="sr-only">加载中...</span></div></div>');

        // 获取当前机台的机型ID
        getCurrentMachineModeId()
            .then(function(modeId) {
                if (!modeId) {
                    console.warn('无法获取当前机台类型');
                    container.empty(); // 清空加载提示，静默处理
                    return;
                }
                
                // 获取同类型的所有机台
                return getMachinesByModeAPI(modeId);
            })
            .then(function(machines) {
                if (!machines || machines.length === 0) {
                    container.html('<div class="col-12 text-center text-muted">没有找到同类型的机台</div>');
                    return;
                }
                
                // 过滤掉当前机台
                const availableMachines = machines.filter(machine => 
                    machine.MachineId.toString() !== currentMachineId.toString()
                );
                
                if (availableMachines.length === 0) {
                    container.html('<div class="col-12 text-center text-muted">没有其他同类型机台可移动</div>');
                    return;
                }
                
                // 生成机台选择卡片（使用与状态卡片相同的逻辑）
                availableMachines.forEach(function(machine) {
                    generateMachineSelectionCardForMove(machine, container);
                });
            })
            .catch(function(error) {
                console.error('获取机台列表失败:', error);
                // 不显示错误提示，静默处理
        });
    }

    // 移动机台到目标机台
    function moveMachineToTarget(orderNumber, targetMachine) {
        const confirmBtn = $('#confirmMoveMachine');
        
        // 显示加载状态
        confirmBtn.addClass('loading').prop('disabled', true).text('移动中...');
        
        // 调用移动机台API
        moveMachineAPI(orderNumber, currentMachineId, targetMachine)
            .then(function(response) {
                // 移动成功
                // 关闭移动机台模态框
                $('#moveMachineModal').modal('hide');
                
                // 立即清理遮罩层
                $('.modal-backdrop').remove();
                $('body').removeClass('modal-open');
                
                // 延迟显示结果模态框，确保移动机台模态框完全关闭
                setTimeout(function() {
                    showResultModal('success', '移动成功', response.message);
                }, 300);
                
                // 更新当前机台数据
                currentMachineData = measurementData[currentMachineId] || [];
                renderTableData(currentMachineId, currentPage);
                
                // 更新机台状态
                updateMachineStatus();
                
                // 清除选择
                $('.row-checkbox').prop('checked', false);
                $('#selectAll').prop('checked', false);
                updateMoveMachineButtonState();
            })
            .catch(function(error) {
                // 移动失败
                showResultModal('error', '移动失败', error.message || '移动机台时发生错误，请稍后重试');
            })
            .finally(function() {
                // 重置按钮状态
                confirmBtn.removeClass('loading').prop('disabled', false).text('移动机台');
            });
    }

    // 批量移动机台到目标机台
    function batchMoveMachinesToTarget(selectedOrders, targetMachine) {
        const confirmBtn = $('#confirmBatchMoveMachine');
        
        // 显示加载状态
        confirmBtn.addClass('loading').prop('disabled', true).text('批量移动中...');
        
        // 创建批量移动的Promise数组
        const movePromises = selectedOrders.map(function(orderNumber) {
            return moveMachineAPI(orderNumber, currentMachineId, targetMachine);
        });
        
        // 执行所有移动操作
        Promise.allSettled(movePromises)
            .then(function(results) {
                // 统计成功和失败的数量
                const successful = results.filter(result => result.status === 'fulfilled').length;
                const failed = results.filter(result => result.status === 'rejected').length;
                
                // 关闭批量移动机台模态框
                $('#batchMoveMachineModal').modal('hide');
                
                // 立即清理遮罩层
                $('.modal-backdrop').remove();
                $('body').removeClass('modal-open');
                
                // 显示批量移动结果
                setTimeout(function() {
                    if (failed === 0) {
                        // 全部成功
                        showResultModal('success', '批量移动成功', 
                            `成功移动 ${successful} 个订单到机台 ${targetMachine}`);
                    } else if (successful === 0) {
                        // 全部失败
                        showResultModal('error', '批量移动失败', 
                            `所有 ${failed} 个订单移动失败，请稍后重试`);
                } else {
                        // 部分成功
                        showResultModal('warning', '批量移动部分成功', 
                            `成功移动 ${successful} 个订单，${failed} 个订单移动失败`);
                    }
                }, 300);
                
                // 更新当前机台数据
                currentMachineData = measurementData[currentMachineId] || [];
                renderTableData(currentMachineId, currentPage);
                
                // 更新机台状态
                updateMachineStatus();
                
                // 清除选择
                $('.row-checkbox').prop('checked', false);
                $('#selectAll').prop('checked', false);
                updateMoveMachineButtonState();
            })
            .catch(function(error) {
                // 批量移动失败
                showResultModal('error', '批量移动失败', error.message || '批量移动机台时发生错误，请稍后重试');
            })
            .finally(function() {
                // 重置按钮状态
                confirmBtn.removeClass('loading').prop('disabled', false).text('批量移动机台');
            });
    }


    // 显示结果模态框
    function showResultModal(type, title, message) {
        const resultModal = $('#resultModal');
        const icon = $('#resultIcon');
        const titleEl = $('#resultTitle');
        const messageEl = $('#resultMessage');
        
        // 设置图标和样式
        icon.empty();
        if (type === 'success') {
            icon.html('<i class="fas fa-check-circle result-icon result-success"></i>');
        } else if (type === 'error') {
            icon.html('<i class="fas fa-times-circle result-icon result-error"></i>');
        } else if (type === 'warning') {
            icon.html('<i class="fas fa-exclamation-triangle result-icon result-warning"></i>');
        } else {
            icon.html('<i class="fas fa-info-circle result-icon result-info"></i>');
        }
        
        titleEl.text(title);
        messageEl.text(message);
        
        // 确保先清理任何现有的模态框实例
        // 显示结果模态框
        $('#resultModal').modal('show');
    }

    function moveSelectedMachines(orders) {
        if (confirm(`确定要移动选中的 ${orders.length} 个订单的机台吗？`)) {
            showAlert(`已成功移动 ${orders.length} 个订单的机台`, 'success');
            // 清除选择
            $('.row-checkbox').prop('checked', false);
            $('#selectAll').prop('checked', false);
        }
    }

    function getSelectedOrders() {
        const selectedOrders = [];
        $('.row-checkbox:checked').each(function() {
            const orderNumber = $(this).closest('tr').find('td:nth-child(2)').text();
            selectedOrders.push(orderNumber);
        });
        console.log('选中的订单:', selectedOrders);
        return selectedOrders;
    }

    function changePage(page) {
        // 实际分页逻辑
        renderTableDataFromCache(currentMachineId, page);
    }

    function refreshData() {
        const refreshBtn = $('#refreshBtn');
        const icon = refreshBtn.find('.fa-sync-alt');
        
        icon.addClass('fa-spin');
        refreshBtn.prop('disabled', true);
        
        setTimeout(function() {
            icon.removeClass('fa-spin');
            refreshBtn.prop('disabled', false);
            showAlert('数据已刷新', 'success');
            // 刷新表格数据
            renderTableData(currentMachineId, currentPage);
            // 更新机台状态
            updateMachineStatus();
        }, 1500);
    }

    function updateMachineStatus() {
        fetchMachineStatusAPI()
            .then(function(response) {
                // 更新机台状态数据
                response.data.forEach(function(machine) {
                    const index = machineStatusData.findIndex(item => item.id === machine.id);
                    if (index > -1) {
                        machineStatusData[index] = machine;
                    }
                });
                // 重新生成机台状态卡片
                generateMachineStatusCards();
            })
            .catch(function(error) {
                console.error('更新机台状态失败:', error);
                // 即使失败也重新生成卡片
                generateMachineStatusCards();
            });
    }

    function showAlert(message, type) {
        // 移除之前的提示框
        $('.custom-alert').remove();
        
        // 根据类型设置图标和颜色
        let icon, bgColor, borderColor, textColor;
        switch(type) {
            case 'success':
                icon = 'fas fa-check-circle';
                bgColor = '#d4edda';
                borderColor = '#c3e6cb';
                textColor = '#155724';
                break;
            case 'error':
                icon = 'fas fa-exclamation-circle';
                bgColor = '#f8d7da';
                borderColor = '#f5c6cb';
                textColor = '#721c24';
                break;
            case 'warning':
                icon = 'fas fa-exclamation-triangle';
                bgColor = '#fff3cd';
                borderColor = '#ffeaa7';
                textColor = '#856404';
                break;
            case 'info':
            default:
                icon = 'fas fa-info-circle';
                bgColor = '#d1ecf1';
                borderColor = '#bee5eb';
                textColor = '#0c5460';
                break;
        }
        
        // 创建美观的提示框
        const alertHtml = `
            <div class="custom-alert" style="
                position: fixed;
                top: 20px;
                right: 20px;
                z-index: 9999;
                min-width: 320px;
                max-width: 400px;
                background: ${bgColor};
                border: 1px solid ${borderColor};
                border-radius: 8px;
                padding: 16px 20px;
                box-shadow: 0 4px 12px rgba(0,0,0,0.15);
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                font-size: 14px;
                line-height: 1.4;
                color: ${textColor};
                transform: translateX(100%);
                transition: transform 0.3s ease-in-out;
                display: flex;
                align-items: center;
                gap: 12px;
            ">
                <i class="${icon}" style="font-size: 18px; flex-shrink: 0;"></i>
                <span style="flex: 1;">${message}</span>
                <button type="button" class="close-btn" style="
                    background: none;
                    border: none;
                    color: ${textColor};
                    font-size: 18px;
                    cursor: pointer;
                    padding: 0;
                    margin-left: 8px;
                    opacity: 0.7;
                    transition: opacity 0.2s;
                ">&times;</button>
            </div>
        `;
        
        $('body').append(alertHtml);
        
        // 显示动画
        setTimeout(function() {
            $('.custom-alert').css('transform', 'translateX(0)');
        }, 10);
        
        // 关闭按钮事件
        $('.custom-alert .close-btn').click(function() {
            hideAlert();
        });
        
        // 3秒后自动关闭
        setTimeout(function() {
            hideAlert();
        }, 3000);
    }

    function hideAlert() {
        $('.custom-alert').css('transform', 'translateX(100%)');
        setTimeout(function() {
            $('.custom-alert').remove();
        }, 300);
    }


    // 键盘快捷键
    $(document).keydown(function(e) {
        // Ctrl + Enter: 添加到队列（打开弹窗）
        if (e.ctrlKey && e.keyCode === 13) {
            if ($('#orderNumber1').is(':focus')) {
                e.preventDefault();
                $('#addToQueueBtn').click();
            }
        }
        
        // F5: 刷新数据
        if (e.keyCode === 116) {
            e.preventDefault();
            refreshData();
        }
    });

    // 输入框回车事件：默认触发“开始/结束量测”
    $('#orderNumber1').keypress(function(e) {
        if (e.which === 13) {
            e.preventDefault();
            $('#toggleMeasurementBtn').click();
        }
    });
    
    // orderNumber2 输入框已被移除，如果有其他输入框需要处理，请单独添加

    // ==================== 量测单号查询功能 ====================
    
    // 根据量测单号查询API
    function searchQcRecordByIdAPI(qcRecordId, companyId) {
        return new Promise(function(resolve, reject) {
            $.ajax({
                url: '/QmsSchedule/SearchQcRecordById',
                type: 'POST',
                data: (function() {
                    const payload = { qcRecordId: qcRecordId };
                    if (companyId && !isNaN(companyId) && companyId > 0) {
                        payload.companyId = companyId;
                    }
                    return payload;
                })(),
                dataType: 'json',
                success: function(response) {
                    if (response.status === 'success') {
                        resolve(response.result);
                    } else {
                        reject(new Error(response.msg || '查询失败'));
                    }
                },
                error: function(xhr, status, error) {
                    reject(new Error('查询API调用失败: ' + error));
                }
            });
        });
    }
    
    // 查询量测单号
    function searchQcRecordById(qcRecordId) {
        // 显示加载状态
        showAlert('正在查询量测单号...', 'info');
        
        searchQcRecordByIdAPI(qcRecordId, currentCompanyId)
            .then(function(qcRecord) {
                console.log('查询到量测单据:', qcRecord);
                
                // 检查量测状态
                if (qcRecord.CheckQcMeasureData === 'F') {
                    showAlert('该量测单已被驳回，无法添加到队列', 'warning');
                    return;
                }
                
                if (qcRecord.CheckQcMeasureData === 'E') {
                    showAlert('该量测单已完成，无法添加到队列', 'warning');
                    return;
                }
                
                // 显示接收/驳回弹窗，并传入完整的量测单据信息
                showReceiveRejectModal(qcRecord);
            })
            .catch(function(error) {
                console.error('查询量测单号失败:', error);
                showAlert('查询失败: ' + error.message, 'error');
            });
    }

    // ==================== 接收/驳回功能 ====================
    
    let currentQcRecord = null;
    
    // 显示接收/驳回弹窗
    function showReceiveRejectModal(qcRecord) {
        currentQcRecord = qcRecord;
        
        // 更新弹窗中的量测单据信息
        $('#modalQcRecordId').text(qcRecord.QcRecordId);
        $('#modalMtlItemNo').text(qcRecord.MtlItemNo);
        $('#modalMtlItemName').text(qcRecord.MtlItemName);
        $('#modalMtlItemSpec').text(qcRecord.MtlItemSpec || '无');
        $('#modalWipNo').text(qcRecord.WipNo);
        $('#modalQcTypeName').text(qcRecord.QcTypeName);
        $('#modalCreateByName').text(qcRecord.CreateByName);
        $('#modalCreateDate').text(new Date(qcRecord.CreateDate).toLocaleString());
        $('#modalStatusName').text(qcRecord.CheckQcMeasureDataName);
        
        // 显示弹窗
        $('#receiveRejectModal').modal('show');
    }
    
    // 接收按钮点击
    $('#receiveBtn').click(function() {
        if (currentQcRecord) {
            receiveQcRecord(currentQcRecord);
        }
    });
    
    // 驳回按钮点击
    $('#rejectBtn').click(function() {
        $('#receiveRejectModal').modal('hide');
        $('#rejectReasonModal').modal('show');
    });
    
    // 确认驳回按钮点击
    $('#confirmRejectBtn').click(function() {
        const reason = $('#rejectReason').val().trim();
        if (!reason) {
            showAlert('请输入驳回原因', 'warning');
            return;
        }
        
        if (currentQcRecord) {
            rejectQcRecord(currentQcRecord, reason);
        }
    });
    
    // 接收送检单
    function receiveQcRecord(qcRecord) {
        const employeeId = sessionStorage.getItem('loginEmployeeId');
        const employeeInfo = employeeData[employeeId];
        const userId = employeeInfo ? employeeInfo.UserId : 1; // 使用真实的UserId
        
        $.ajax({
            url: '/QmsSchedule/AddToSchedule',
            type: 'POST',
            data: {
                qcRecordId: qcRecord.QcRecordId,
                machineId: currentMachineId,
                userId: userId
            },
            dataType: 'json',
            success: function(response) {
                if (response.status === 'success') {
                    // 关闭接收/驳回弹窗
                    $('#receiveRejectModal').modal('hide');
                    
                    // 清理弹窗内容
                    currentQcRecord = null;
                    
                    // 显示成功提示
                    showAlert('送检单接收成功', 'success');
                    
                    // 刷新数据
                    loadMachineData(currentMachineId);
                } else {
                    showAlert('接收失败: ' + response.msg, 'error');
                }
            },
            error: function(xhr, status, error) {
                showAlert('网络错误: ' + error, 'error');
            }
        });
    }
    
    // 驳回送检单
    function rejectQcRecord(qcRecord, reason) {
        const employeeId = sessionStorage.getItem('loginEmployeeId');
        const employeeInfo = employeeData[employeeId];
        const userId = employeeInfo ? employeeInfo.UserId : 1; // 使用真实的UserId
        
        $.ajax({
            url: '/QmsSchedule/RejectQcRecord',
            type: 'POST',
            data: {
                qcRecordId: qcRecord.QcRecordId,
                reason: reason,
                userId: userId
            },
            dataType: 'json',
            success: function(response) {
                if (response.status === 'success') {
                    // 关闭驳回原因弹窗
                    $('#rejectReasonModal').modal('hide');
                    
                    // 清空驳回原因输入框
                    $('#rejectReason').val('');
                    
                    // 清理弹窗内容
                    currentQcRecord = null;
                    
                    // 显示成功提示
                    showAlert('送检单驳回成功', 'success');
                } else {
                    showAlert('驳回失败: ' + response.msg, 'error');
                }
            },
            error: function(xhr, status, error) {
                showAlert('网络错误: ' + error, 'error');
            }
        });
    }

    // 清空表格数据
    function clearTableData() {
        const tbody = $('#orderTable tbody');
        tbody.empty();
        
        // 清空分页信息
        $('#paginationInfo').text('共 0 条记录');
        $('#pagination').empty();
        
        // 清空当前数据缓存
        currentMachineData = [];
        currentPage = 1;
        totalPages = 1;
        
        console.log('表格数据已清空');
    }
    
    // 清空机台状态
    function clearMachineStatus() {
        // 清空机台状态显示
        $('#machineStatus').text('未知');
        $('#machineStatus').removeClass('status-queue status-measuring status-completed');
        
        // 清空机台信息
        $('#currentMachineId').text('未知机台');
        
        console.log('机台状态已清空');
    }

    // 初始化页面数据
    initializePageData();

    console.log('量测排程系统已初始化');
});
