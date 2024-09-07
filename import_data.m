% *******************************本项目基于MATLAB实现EIS数据的批量DRT变换处理*********************************
% 使用说明
% 作者：Lucheng Cai，浙江大学材料科学与工程学院Wei-Qiang Han老师课题组直博生
% 1.使用DRT拟合可修改以下参数（line55-line61）
% 2.本代码默认使用贝叶斯法拟合（Bayesian Run），如不能使用此法，将使用简单拟合（Simple Run）。具体可见代码运行完毕后的提示。
% 3.用户需提供包含EIS数据的excel表，每个EIS数据包含‘频率’‘阻抗实部’‘阻抗虚部’分列保存，不同EIS数据放在excel表不同sheet中，sheet命名应符合规则（尽量以英文字母和数字组合）。具体见EIS数据模板。
% 4.用户需根据EIS数据中‘频率’‘阻抗实部’‘阻抗虚部’的位置修改line97-line99的代码，具体见该行相应注释。

% 使用方法
% 1.新建项目文件夹，下载本项目所有文件（后缀.m）到项目文件夹
% 2.将多个EIS数据放入到项目文件夹中的excel表（新建）中，具体要求见上述 使用说明3.
% 3.打开matlab，将matlab左侧文件夹设置为项目文件夹，打开本项目中的import_data.m文件
% 4.修改拟合参数（line55-line61），一般默认即可。
% 5.关键！关键！关键！根据EIS数据中‘频率’‘阻抗实部’‘阻抗虚部’的位置修改line97-line99的代码，具体见该行相应注释。
% 6.点击matlab上方运行按钮，选定目标excel表，等待代码运行完毕。
% 7.关键！关键！关键！作图。在生成的totalDRT的excel表中，每两列为一组EIS数据，各组的前一列为tau，后一列为MAP。整体复制到origin中，前者为横轴，后者为纵轴即可作图。

% 本项目在香港科技大学Francesco Ciucci老师工作的基础上进行改进，具体请见：https://github.com/ciuccislab/DRTtools
% 在此对Francesco Ciucci老师抱以敬意！

% 感谢好朋友Chenchen Xu的指导解惑（github：https://github.com/InvincibleGuy777）

% *******************************Batch DRT transformation processing of EIS data based on MATLAB*********************************
% Instructions:
% Author: Lucheng Cai, a doctoral candidate in the research group of Professor Wei-Qiang Han at the School of Materials Science and Engineering, Zhejiang University.
% 1.Parameters for DRT fitting can be modified (line 55-line 61).
% 2.This code defaults to using Bayesian Run for fitting. If this method cannot be used, a Simple Run fitting will be applied. Details can be seen in the prompt after running the code.
% 3.Users need to provide an Excel containing EIS data. Each EIS data should include columns for 'Frequency,' 'real part of impedance (Re),' and 'imaginary part of impedance (Im).' 
%   Different EIS data should be placed in different sheets of the Excel, and the sheet names should follow the naming convention (preferably a combination of letters and numbers). 
%   Refer to the EIS data template for specifics.
% 4.Users should modify the code at lines 97-99 according to the positions of 'Frequency,' 'Re,' and 'Im' in the EIS data.

% Directions:
% 1.Create a new project folder and download all files of this project (with the .m extension) into the project folder.
% 2.Place multiple EIS data into an Excel (create a new one) in the project folder, following the requirements mentioned in the Instructions above.
% 3.Open Matlab, set the folder on the left side to the project folder, and open the import_data.m file in this project.
% 4.Modify the fitting parameters (line 55-line 61), usually the default settings are sufficient.
% 5.Crucial! Crucial! Crucial! Modify the code at lines 97-99 according to the positions of 'Frequency,' 'Re,' and 'Im' in the EIS data. 
%   Refer to the corresponding comments in those lines for specifics.
% 6.Click the run button at the top of Matlab, select the target Excel table, and wait for the code to finish running.
% 7.Crucial! Crucial! Crucial! Plotting. In the generated total DRT Excel table, each two columns represent one set of EIS data, where the first column is tau and the second column is MAP. 
%   Copy the data to another sheet, using the former as the x-axis and the latter as the y-axis for plotting.

% This project is an improvement based on the work of Professor Francesco Ciucci at the Hong Kong University of Science and Technology. 
% For more details, please visit: https://github.com/ciuccislab/DRTtools
% A respectful acknowledgement to Professor Francesco Ciucci!

% Thanks to my friend Chenchen Xu for guidance and assistance (github: https://github.com/InvincibleGuy777)



function import_data(handles)
 clc;
  
    handles.rbf_type = 'Gaussian';   %Method of Discretization：‘Gaussian’(默认)，‘C2 Metern’，‘C4 Metern’，‘C6 Metern’，‘Inverse quadratic’，‘Inverse quadric’，'Cauthy','Piecewise linear'
    handles.data_used = 'Combined Re-Im Data';  %Data used:'Combined Re-Im Data'，‘Re Data’，‘Im Data’
    handles.lambda = 1e-3;  %Regularization parameter
    handles.sample_number = 2000;  %Number of Samples(for Bayesian Run)
    handles.coeff = 0.5;  %FWHM control or Shape factor
    handles.shape_control = 'FWHM Coefficient';   %RBF shape control
    handles.der_used = '1st-order';   %Regularization derivative:'1st-order',‘2nd-order’
    
    strArraySim = ["Simple Run: "];
    strArrayBay = ["Bayesian Run: "];
% 选择要打开的Excel文件
[fileName, filePath] = uigetfile('*.xlsx', 'Select an Excel file');
folder = filePath;
if fileName == 0
    disp('用户取消了操作');
else
    % 读取Excel文件
    [~, sheetNames] = xlsfinfo(fullfile(filePath, fileName));
    % 遍历每个工作表
    for sheetIndex = 1:numel(sheetNames)
        try
            sheetName = sheetNames{sheetIndex};
        catch ME
            disp('***************************************************');
            disp('*************请重启MATLAB再次运行本代码************');
            disp('***Please restart MATLAB and run this code again***');
            disp('*************请重启MATLAB再次运行本代码************');
            disp('***Please restart MATLAB and run this code again***');
            disp('***************************************************');
        end
        
        sheetName = sheetNames{sheetIndex};
        txtFileName = fullfile(filePath, [sheetName, '.txt']); % 使用工作表名字来命名txt文件，并保存到与Excel文件相同的位置

        % 读取数据
        data = xlsread(fullfile(filePath, fileName), sheetName);

        % 移除列标题行
        data = data(2:end, :);

        % 获取所需的数据列
        % EIS数据中第一行为各列的名称
        columnA = data(:, 1); % A列数据   ‘频率’的位置，在A列参数为1，在B列参数需改为2，以此类推，笔者此处为A列。
        columnE = data(:, 5); % E列数据   ‘阻抗实部’的位置，在A列参数为1，在B列参数需改为2，以此类推，笔者此处为E列。
        columnF = data(:, 6); % F列数据   ‘阻抗虚部’的位置，在A列参数为1，在B列参数需改为2，以此类推，笔者此处为F列。

        % 将数据写入txt文件
        fid = fopen(txtFileName, 'w');
        fprintf(fid, '%f\t%f\t%f\n', [columnA, columnE, columnF]');
        fclose(fid);

        disp(['工作表 "', sheetName, '" 的数据已导出到 ', txtFileName]);
    end

    disp('所有工作表的数据已成功导出为txt文件。');
end
   
%   method_tag: 'none': havnt done any computation, 'simple': simple DRT,
%               'credit': Bayesian run, 'BHT': Bayesian Hibert run
    handles.method_tag = 'none'; 
    handles.data_exist = false;
    startingFolder = 'C:\*';
    if ~exist(startingFolder, 'dir')
%       if that folder doesn't exist, just start in the current folder.
        startingFolder = pwd;
    end

    
    txtFiles = dir(fullfile(filePath, '*.txt')); % 获取文件夹中所有txt文件的信息
    numFiles = numel(txtFiles);
    
    
    if numFiles == 0
        disp('文件夹中没有txt文件');
    else
        for i = 1:numFiles
            baseFileName = txtFiles(i).name; % 获取当前文件名
            fullFileName = fullfile(folder, baseFileName); % 获取当前文件的完整路径

    dealname = baseFileName;
    dealfolder = folder
    [folder,baseFileName,ext] = fileparts(fullFileName);

    
    
    
    if ~baseFileName
%       User clicked the Cancel button.
        return;
    end

    switch ext
        case '.mat' % User selects Mat files.
            storedStructure = load(fullFileName);
            A = [storedStructure.freq,storedStructure.Z_prime,storedStructure.Z_double_prime];
            
            handles.data_exist = true;

        case '.txt' % User selects Txt files.
        %   change comma to dot if necessary
            fid  = fopen(fullFileName,'r');
            f1 = fread(fid,'*char')';
            fclose(fid);

            baseFileName = strrep(f1,',','.');
            fid  = fopen(fullFileName,'w');
            fprintf(fid,'%s',baseFileName);
            fclose(fid);
            try
                % 可能会报错的代码块
                A = dlmread(fullFileName);
            catch ME
                % 异常处理代码块
                disp('***********************************************************************');
                disp('****************请将文件夹中所有.txt文件删除后再运行代码***************');
                disp('***Please delete all.txt files in the folder before running the code***');
                disp('****************请将文件夹中所有.txt文件删除后再运行代码***************');
                disp('***Please delete all.txt files in the folder before running the code***');
                disp('***********************************************************************');
                disp(ME.message);
            end
            
            A = dlmread(fullFileName);
            
        %   change back dot to comma if necessary    
            fid  = fopen(fullFileName,'w');
            fprintf(fid,'%s',f1);
            fclose(fid);
            
            handles.data_exist = true;

        case '.csv' % User selects csv.
            A = csvread(fullFileName);
            
            handles.data_exist = true;

        otherwise
            warning('Invalid file type')
            handles.data_exist = false;
    end

%   find incorrect rows with zero frequency
    index = find(A(:,1)==0); 
    A(index,:)=[];
    
%   flip freq, Z_prime and Z_double_prime so that data are in the desceding 
%   order of freq 
    if A(1,1) < A(end,1)
       A = fliplr(A')';
    end
    
    handles.freq = A(:,1);
    handles.Z_prime_mat = A(:,2);
    handles.Z_double_prime_mat = A(:,3);
    
%   save original freq, Z_prime and Z_double_prime
    handles.freq_0 = handles.freq;
    handles.Z_prime_mat_0 = handles.Z_prime_mat;
    handles.Z_double_prime_mat_0 = handles.Z_double_prime_mat;
    
    handles.Z_exp = handles.Z_prime_mat(:)+ 1i*handles.Z_double_prime_mat(:);
    
    handles.method_tag = 'none';
   
    
    %开始处理
    
    if ~handles.data_exist
        return
    end    


%   bounds ridge regression
    handles.lb = zeros(numel(handles.freq)+2,1);
    handles.ub = Inf*ones(numel(handles.freq)+2,1);
    handles.x_0 = ones(size(handles.lb));

    handles.options = optimset('algorithm','interior-point-convex','Display','off','TolFun',1e-15,'TolX',1e-10,'MaxFunEvals', 1E5);

    handles.b_re = real(handles.Z_exp);
    handles.b_im = imag(handles.Z_exp);

%   compute epsilon
    handles.epsilon = compute_epsilon(handles.freq, 0.5, 'Gaussian', 'FWHM Coefficient');

%   calculate the A_matrix
    handles.A_re = assemble_A_re(handles.freq, handles.epsilon, 'Gaussian');
    handles.A_im = assemble_A_im(handles.freq, handles.epsilon, 'Gaussian');

%   adding the resistence column to the A_re_matrix
    handles.A_re(:,2) = 1;
    
%   adding the inductance column to the A_im_matrix if necessary
%    if  get(handles.inductance,'Value')==2
%        handles.A_im(:,1) = 2*pi*(handles.freq(:));
%    end
    
%   calculate the M_matrix
    switch handles.der_used
        case '1st-order'
            handles.M = assemble_M_1(handles.freq, handles.epsilon, handles.rbf_type);
        case '2nd-order'
            handles.M = assemble_M_2(handles.freq, handles.epsilon, handles.rbf_type);
    end

%   Running ridge regression
    switch handles.data_used
        case 'Combined Re-Im Data'
            [H_combined,f_combined] = quad_format_combined(handles.A_re, handles.A_im, handles.b_re, handles.b_im, handles.M, handles.lambda);
            handles.x_ridge = quadprog(H_combined, f_combined, [], [], [], [], handles.lb, handles.ub, handles.x_0, handles.options);

            %prepare for HMC sampler
            handles.mu_Z_re = handles.A_re*handles.x_ridge;
            handles.mu_Z_im = handles.A_im*handles.x_ridge;

            handles.res_re = handles.mu_Z_re-handles.b_re;
            handles.res_im = handles.mu_Z_im-handles.b_im;

            sigma_re_im = std([handles.res_re;handles.res_im]);

            inv_V = 1/sigma_re_im^2*eye(numel(handles.freq));

            Sigma_inv = (handles.A_re'*inv_V*handles.A_re) + (handles.A_im'*inv_V*handles.A_im) + (handles.lambda/sigma_re_im^2)*handles.M;
            mu_numerator = handles.A_re'*inv_V*handles.b_re + handles.A_im'*inv_V*handles.b_im;
            
        case 'Im Data'
            [H_im,f_im] = quad_format(handles.A_im, handles.b_im, handles.M, handles.lambda);
            handles.x_ridge = quadprog(H_im, f_im, [], [], [], [], handles.lb, handles.ub, handles.x_0, handles.options);

            %prepare for HMC sampler
            handles.mu_Z_re = handles.A_re*handles.x_ridge;
            handles.mu_Z_im = handles.A_im*handles.x_ridge;

            handles.res_im = handles.mu_Z_im-handles.b_im;
            sigma_re_im = std(handles.res_im);

            inv_V = 1/sigma_re_im^2*eye(numel(handles.freq));

            Sigma_inv = (handles.A_im'*inv_V*handles.A_im) + (handles.lambda/sigma_re_im^2)*handles.M;
            mu_numerator = handles.A_im'*inv_V*handles.b_im;

        case 'Re Data'
            [H_re,f_re] = quad_format(handles.A_re, handles.b_re, handles.M, handles.lambda);
            handles.x_ridge = quadprog(H_re, f_re, [], [], [], [], handles.lb, handles.ub, handles.x_0, handles.options);

            %prepare for HMC sampler
            handles.mu_Z_re = handles.A_re*handles.x_ridge;
            handles.mu_Z_im = handles.A_im*handles.x_ridge;

            handles.res_re = handles.mu_Z_re-handles.b_re;
            sigma_re_im = std(handles.res_re);

            inv_V = 1/sigma_re_im^2*eye(numel(handles.freq));

            Sigma_inv = (handles.A_re'*inv_V*handles.A_re) + (handles.lambda/sigma_re_im^2)*handles.M;            
            mu_numerator = handles.A_re'*inv_V*handles.b_re;

    end
    
    warning('off','MATLAB:singularMatrix')
    handles.Sigma_inv = (Sigma_inv+Sigma_inv')/2;
    handles.mu = handles.Sigma_inv\mu_numerator; % linsolve
    
%   method_tag: 'none': havnt done any computation, 'simple': simple DRT,
%               'credit': Bayesian run, 'BHT': Bayesian Hilbert run
    handles.method_tag = 'simple'; 



    %贝叶斯

%   Running HMC sampler
    handles.mu = handles.mu(3:end);
    handles.Sigma_inv = handles.Sigma_inv(3:end,3:end);
    handles.Sigma = inv(handles.Sigma_inv);

    F = eye(numel(handles.x_ridge(3:end)));
    g = eps*ones(size(handles.x_ridge(3:end)));
    initial_X = handles.x_ridge(3:end)+100*eps;
    
    
    handles.Xs = HMC_exact(F, g, handles.Sigma, handles.mu, true, handles.sample_number, initial_X);
        % handles.lower_bound = quantile(handles.Xs(:,500:end),.005,2);
        % handles.upper_bound = quantile(handles.Xs(:,500:end),.995,2);
        handles.lower_bound = quantile_alter(handles.Xs(:,500:end),.005,2,'R-5');
        handles.upper_bound = quantile_alter(handles.Xs(:,500:end),.995,2,'R-5');
        handles.mean = mean(handles.Xs(:,500:end),2);
 
   %plot贝叶斯
   taumax = ceil(max(log10(1./handles.freq)))+0.5;    
   taumin = floor(min(log10(1./handles.freq)))-0.5;
   handles.freq_fine = logspace(-taumin, -taumax, 10*numel(handles.freq));
   
   
    if ~any(isnan(handles.mean))
        strArrayBay = [strArrayBay, dealname];
    
   [handles.gamma_mean_fine,handles.freq_fine] = map_array_to_gamma(handles.freq_fine, handles.freq, handles.mean, handles.epsilon, handles.rbf_type);
   [handles.lower_bound_fine,handles.freq_fine] = map_array_to_gamma(handles.freq_fine, handles.freq, handles.lower_bound, handles.epsilon, handles.rbf_type);
   [handles.upper_bound_fine,handles.freq_fine] = map_array_to_gamma(handles.freq_fine, handles.freq, handles.upper_bound, handles.epsilon, handles.rbf_type);
   [handles.gamma_ridge_fine,handles.freq_fine] = map_array_to_gamma(handles.freq_fine, handles.freq, handles.x_ridge(3:end), handles.epsilon, handles.rbf_type);
    
    % 指定文件名
    fileName = ['drt' dealname];

% 构建完整的文件路径
    fullFileName = fullfile(dealfolder, fileName);

% 使用 'w+' 模式打开文件以供读写
    fid = fopen(fullFileName, 'wt');
    col_tau = 1./handles.freq_fine(:);
    col_gamma = handles.gamma_ridge_fine(:);
    col_mean = handles.gamma_mean_fine(:);
    col_upper = handles.upper_bound_fine(:);
    col_lower = handles.lower_bound_fine(:);
    
    fprintf(fid,'%s, %e \n','L',handles.x_ridge(1));
    fprintf(fid,'%s, %e \n','R',handles.x_ridge(2));
    fprintf(fid,'%s, %s, %s, %s, %s\n', 'tau', 'MAP', 'Mean', 'Upperbound', 'Lowerbound');
    fprintf(fid,'%e, %e, %e, %e, %e\n', [col_tau(:), col_gamma(:), col_mean(:), col_upper(:), col_lower(:)]');
    
    fclose(fid);
   
    else
 % 指定文件名    报错的关键
    fileName = ['drt' dealname];
    strArraySim = [strArraySim, dealname];

% 构建完整的文件路径
    fullFileName = fullfile(dealfolder, fileName);

% 使用 'w+' 模式打开文件以供读写
    fid = fopen(fullFileName, 'wt');
    
    [handles.gamma_ridge_fine,handles.freq_fine] = map_array_to_gamma(handles.freq_fine, handles.freq, handles.x_ridge(3:end), handles.epsilon, handles.rbf_type);
    col_tau = 1./handles.freq_fine(:);
    col_gamma = handles.gamma_ridge_fine(:);
    fprintf(fid,'%s, %e \n','L',handles.x_ridge(1));
    fprintf(fid,'%s, %e \n','R',handles.x_ridge(2));
    fprintf(fid,'%s, %s \n','tau','gamma(tau)');
    fprintf(fid,'%e, %e \n', [col_tau(:), col_gamma(:)]');
    
    fclose(fid);
        
     end
    end
    end
    


% 文件夹路径
path = filePath;
% 保存数据为Excel文件
outputFile = fullfile(path, 'totalDRT.xlsx');

% 获取文件夹中的所有.txt文件
fileList = dir(fullfile(path, 'drt*.txt'));

% 创建一个空的矩阵用于存储数据
data = [];
%起始写入列的位置
startColumn = 'A';
startAscii = double(startColumn);

% 遍历文件列表
for i = 1:numel(fileList)
    % 读取当前文件
    filename = fullfile(path, fileList(i).name);
    fileData = readmatrix(filename, 'Delimiter', ',');
    
    cleanName = strrep(fileList(i).name, 'drt', '');
    %{
    disp(fileList(i).name);
    disp(strArraySim);
    
    disp(cleanName);
   %}
    
    % 提取第1、2列数据
    columns = fileData(:, 1:2);
    
    if ismember(cleanName, strArraySim)
        columns(1:3, :) = [];
    end

    disp(size(columns))
    
    %写入第一行文件名
    location = [startColumn,'1'];
    writematrix([filename], outputFile, 'Range', location);
    
 
    %写入数据
    location = [startColumn,'2'];
    writematrix(columns, outputFile, 'Range', location);

    %更新列位置
    startAscii = double(startColumn);
    startColumn = char(startAscii+2);
    
end

disp(strArraySim)
disp(strArrayBay)
disp('***Program complete***')

