function varargout = batch_curve_processor(varargin)
% BATCH_CURVE_PROCESSOR MATLAB code for batch_curve_processor.fig
%      BATCH_CURVE_PROCESSOR, by itself, creates a new BATCH_CURVE_PROCESSOR or raises the existing
%      singleton*.
%
%      H = BATCH_CURVE_PROCESSOR returns the handle to a new BATCH_CURVE_PROCESSOR or the handle to
%      the existing singleton*.
%
%      BATCH_CURVE_PROCESSOR('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in BATCH_CURVE_PROCESSOR.M with the given input arguments.
%
%      BATCH_CURVE_PROCESSOR('Property','Value',...) creates a new BATCH_CURVE_PROCESSOR or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before batch_curve_processor_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to batch_curve_processor_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help batch_curve_processor

% Last Modified by GUIDE v2.5 07-Sep-2025 22:13:12

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @batch_curve_processor_OpeningFcn, ...
                   'gui_OutputFcn',  @batch_curve_processor_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before batch_curve_processor is made visible.
function batch_curve_processor_OpeningFcn(hObject, eventdata, handles, varargin)
handles.output = hObject;

% --- 初始化我们的变量 ---
handles.all_data = {};            % Cell数组，用于存储所有曲线数据
handles.num_curves = 0;           % 曲线总数
handles.current_curve_index = 0;  % 当前显示的曲线索引
handles.results = struct();       % 用于存储最终结果的结构体数组

% 初始化按钮状态，加载数据前禁用大部分按钮
set(handles.prev_button, 'Enable', 'off');
set(handles.next_button, 'Enable', 'off');
set(handles.manual_calc_button, 'Enable', 'off');
set(handles.mark_abnormal_button, 'Enable', 'off');
set(handles.save_button, 'Enable', 'off');
% --- 初始化结束 ---

guidata(hObject, handles);

% --- Outputs from this function are returned to the command line.
function varargout = batch_curve_processor_OutputFcn(hObject, eventdata, handles) 
varargout{1} = handles.output;


% --- Executes on button press in load_button.
function load_button_Callback(hObject, eventdata, handles)
[file, path] = uigetfile('*.xlsx;*.xls', '请选择包含曲线数据的Excel文件');
if isequal(file, 0)
    set(handles.status_text, 'String', '状态：用户取消了选择');
    return;
end

fullFilePath = fullfile(path, file);
try
    set(handles.status_text, 'String', '状态：正在读取数据，请稍候...');
    drawnow; % 强制刷新UI
    
    % --- 使用经典的 xlsread 函数 ---
    % xlsread 会返回一个纯数字的矩阵
    [numeric_data, ~, ~] = xlsread(fullFilePath);
    
    if isempty(numeric_data)
        errordlg('Excel文件中没有找到任何数字数据，请检查文件。', '读取错误');
        set(handles.status_text, 'String', '状态：文件内容为空或格式错误');
        return;
    end
    
    [~, num_cols] = size(numeric_data);
    
    % 检查列数是否为偶数
    if mod(num_cols, 2) ~= 0
        errordlg('Excel文件中的数字列数不是偶数，请检查数据格式。', '格式错误');
        set(handles.status_text, 'String', '状态：文件格式错误');
        return;
    end
    
    % 解析数据，每两列为一条曲线
    handles.num_curves = num_cols / 2;
    handles.all_data = cell(1, handles.num_curves);
    
    % --- 初始化结果结构体数组 ---
    handles.results = repmat(struct(...
        'CurveNumber', 0, ...
        'AvgLoad_5_45mm', NaN, ...
        'AvgLoad_20_30mm', NaN, ...
        'ManualRangeStart', NaN, ...
        'ManualRangeEnd', NaN, ...
        'ManualAvgLoad', NaN, ...
        'IsAbnormal', false), handles.num_curves, 1);
        
    for i = 1:handles.num_curves
        % 第 2*i-1 列是荷重 (Y)，第 2*i 列是位移 (X)
        load_col = numeric_data(:, 2*i - 1);
        disp_col = numeric_data(:, 2*i);
        
        % 移除NaN数据对
        valid_idx = ~isnan(load_col) & ~isnan(disp_col);
        handles.all_data{i} = [disp_col(valid_idx), load_col(valid_idx)];
        
        % 填充曲线编号
        handles.results(i).CurveNumber = i;
    end
    
    % 设置当前曲线为第一条并显示
    handles.current_curve_index = 1;
    handles = update_display(handles); % 调用一个辅助函数来更新显示
    
    % 更新按钮状态
    set(handles.prev_button, 'Enable', 'on');
    set(handles.next_button, 'Enable', 'on');
    set(handles.manual_calc_button, 'Enable', 'on');
    set(handles.mark_abnormal_button, 'Enable', 'on');
    set(handles.save_button, 'Enable', 'on');
    
    set(handles.status_text, 'String', ['状态：成功加载 ', num2str(handles.num_curves), ' 条曲线']);
    
catch ME
    errordlg(['读取或解析文件时出错: ' ME.message], '错误');
    set(handles.status_text, 'String', '状态：文件读取失败');
    return;
end

guidata(hObject, handles);


% --- Executes on button press in prev_button.
function prev_button_Callback(hObject, eventdata, handles)
if handles.current_curve_index > 1
    handles.current_curve_index = handles.current_curve_index - 1;
    handles = update_display(handles);
    guidata(hObject, handles);
end


% --- Executes on button press in next_button.
function next_button_Callback(hObject, eventdata, handles)
if handles.current_curve_index < handles.num_curves
    handles.current_curve_index = handles.current_curve_index + 1;
    handles = update_display(handles);
    guidata(hObject, handles);
end


% --- Executes on button press in manual_calc_button.
function manual_calc_button_Callback(hObject, eventdata, handles)
set(handles.status_text, 'String', '状态：请在图上点击两点以确定范围');

try
    [x_coords, ~] = ginput(2);
catch
    set(handles.status_text, 'String', '状态：选择被中断，请重试');
    return;
end

x_start = min(x_coords);
x_end = max(x_coords);

idx = handles.current_curve_index;
current_data = handles.all_data{idx};
displacement = current_data(:, 1);
load_data = current_data(:, 2);

range_idx = displacement >= x_start & displacement <= x_end;
avg_load = mean(load_data(range_idx));

if isnan(avg_load) || isempty(avg_load)
    avg_load = 0; % 如果选区内没点，则为0
end

% 保存结果到结果结构体
handles.results(idx).ManualRangeStart = x_start;
handles.results(idx).ManualRangeEnd = x_end;
handles.results(idx).ManualAvgLoad = avg_load;

% 更新UI
set(handles.manual_calc_text, 'String', sprintf('范围 [%.2f, %.2f], 平均荷重: %.4f', x_start, x_end, avg_load));
set(handles.status_text, 'String', '状态：手动计算完成');

guidata(hObject, handles);


% --- Executes on button press in mark_abnormal_button.
function mark_abnormal_button_Callback(hObject, eventdata, handles)
idx = handles.current_curve_index;

% 切换异常状态
current_state = handles.results(idx).IsAbnormal;
handles.results(idx).IsAbnormal = ~current_state;

% 更新显示
handles = update_display(handles);

guidata(hObject, handles);


% --- Executes on button press in save_button.
function save_button_Callback(hObject, eventdata, handles)
[file, path] = uiputfile('*.xlsx', '保存分析结果', 'processed_results.xlsx');

if isequal(file, 0)
    set(handles.status_text, 'String', '状态：用户取消了保存');
    return;
end

fullSavePath = fullfile(path, file);

try
    % 在保存前，将结构体数组转换为table
    results_table = struct2table(handles.results);
    writetable(results_table, fullSavePath);
    msgbox(['所有结果已成功保存到: ' fullSavePath], '保存成功');
    set(handles.status_text, 'String', '状态：结果已保存');
catch ME
    % 如果writetable失败 (也可能是版本太旧)，提供备用方案
    if strcmp(ME.identifier, 'MATLAB:UndefinedFunction')
        try
            % 尝试使用更旧的 xlswrite
            warning('off', 'MATLAB:xlswrite:AddSheet');
            results_cell = [fieldnames(handles.results)'; struct2cell(handles.results)'];
            xlswrite(fullSavePath, results_cell);
            msgbox(['所有结果已成功保存到: ' fullSavePath], '保存成功 (兼容模式)');
            set(handles.status_text, 'String', '状态：结果已保存');
        catch ME_xlswrite
            errordlg(['保存文件失败，两种方法均告失败: ' ME_xlswrite.message], '保存错误');
            set(handles.status_text, 'String', '状态：保存失败');
        end
    else
       errordlg(['保存文件失败: ' ME.message], '保存错误');
       set(handles.status_text, 'String', '状态：保存失败');
    end
end


% --- 辅助函数：更新图形界面显示 ---
function handles = update_display(handles)
if handles.current_curve_index == 0
    return; % 如果没有数据，则不执行任何操作
end

% 获取当前曲线数据
idx = handles.current_curve_index;
current_data = handles.all_data{idx};
displacement = current_data(:, 1);
load_data = current_data(:, 2);

% 绘图
plot(handles.axes1, displacement, load_data, '-o', 'MarkerSize', 3);
xlabel(handles.axes1, '位移 (mm)');
ylabel(handles.axes1, '荷重');
title(handles.axes1, ['曲线 ', num2str(idx)]);
grid(handles.axes1, 'on');

% --- 自动计算 ---
% 5-45mm
range_idx1 = displacement >= 5 & displacement <= 45;
avg1 = mean(load_data(range_idx1));
if isempty(avg1) || isnan(avg1), avg1 = 0; end % 处理空选区
handles.results(idx).AvgLoad_5_45mm = avg1;

% 20-30mm
range_idx2 = displacement >= 20 & displacement <= 30;
avg2 = mean(load_data(range_idx2));
if isempty(avg2) || isnan(avg2), avg2 = 0; end % 处理空选区
handles.results(idx).AvgLoad_20_30mm = avg2;

% --- 更新UI文本 ---
set(handles.curve_info_text, 'String', ['曲线: ', num2str(idx), ' / ', num2str(handles.num_curves)]);
set(handles.auto_calc1_text, 'String', sprintf('5-45mm 平均荷重: %.4f', avg1));
set(handles.auto_calc2_text, 'String', sprintf('20-30mm 平均荷重: %.4f', avg2));

% 检查并显示之前的手动计算结果和异常状态
manual_start = handles.results(idx).ManualRangeStart;
manual_end = handles.results(idx).ManualRangeEnd;
manual_avg = handles.results(idx).ManualAvgLoad;
if ~isnan(manual_avg)
    set(handles.manual_calc_text, 'String', sprintf('范围 [%.2f, %.2f], 平均荷重: %.4f', manual_start, manual_end, manual_avg));
else
    set(handles.manual_calc_text, 'String', '范围 [-,-], 平均荷重: -');
end

if handles.results(idx).IsAbnormal
    set(handles.status_text, 'String', ['状态：曲线 ', num2str(idx), ' (已标记为异常)']);
    set(handles.mark_abnormal_button, 'BackgroundColor', [1, 0.6, 0.6]); % 按钮变红
else
    set(handles.status_text, 'String', ['状态：当前为曲线 ', num2str(idx)]);
    set(handles.mark_abnormal_button, 'BackgroundColor', get(0,'defaultUicontrolBackgroundColor')); % 按钮恢复默认色
end

% 更新导航按钮的可用状态
set(handles.prev_button, 'Enable', 'on');
set(handles.next_button, 'Enable', 'on');
if idx == 1
    set(handles.prev_button, 'Enable', 'off');
end
if idx == handles.num_curves
    set(handles.next_button, 'Enable', 'off');
end
