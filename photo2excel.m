function photo2excel(filename)
% 利用MATLAB通过填色的方式把图像画入Excel文档(玩转Excel拼图)
%   photo2excel(filename)  输入参数filename为包含图片文件所在路径的字符串
%   photo2excel(I)         输入参数I为图像灰度值矩阵
%   
%   注：图像可以是灰度图像或真彩图像
%   Copyright 2010 - 2011 xiezhh. 
%   $Revision: 1.0.0.0 [        DISCUZ_CODE_1        ]nbsp; $Date: 2011/03/12 13:41:00 $

 

% 判断输入参数类型，导入图像灰度值矩阵
if ischar(filename) && exist(filename,'file');
    I = imread(filename);
elseif isnumeric(filename)
    I = filename;
else
    error('输入参数类型不匹配');
end
I = uint8(I);

 

% 判断图像类型
[m,n,k] = size(I);
if k == 1
    I = repmat(I,[1 1 3]);
end
if k ~= 1 && k ~= 3
    error('图像类型不匹配，只能是灰度图像或真彩图像');
end

 

% 设定要保存的Excel文件名和路径
filespec_user = [pwd '\我的照片.xls'];

 

% 判断Excel是否已经打开，若已打开，就在打开的Excel中进行操作，否则就打开Excel
try
    % 若Excel服务器已经打开，返回其句柄Excel
    Excel = actxGetRunningServer('Excel.Application');
catch
    % 创建一个Microsoft Excel服务器，返回句柄Excel
    Excel = actxserver('Excel.Application'); 
end;

 

% 设置Excel服务器为可见状态
Excel.Visible = 1;

 

% 新建一个工作簿，并保存，文件名为我的照片.Excel
Workbook = Excel.Workbooks.Add;
%Workbook.SaveAs(filespec_user);

 

% 返回当前工作表句柄
Sheets = Excel.ActiveWorkbook.Sheets;
Sheet1 = Sheets.Item(1);    % 返回第1个表格句柄

 

% 设置行高和列宽
SC1 = Sheet1.Cells;
% 设置行高
SC1.RowHeight = 1.5;
% 设置列宽
SC1.ColumnWidth = 0.15;

 

% 通过填色的方式绘图
ColNum = SC1.Columns.Count;
RowNum = SC1.Rows.Count;
I = double(I);
ColorIndex = I(:,:,1) + 256*I(:,:,2) + 65536*I(:,:,3);
for j = 1:n
    for i = 1:m
        SC1.Item((i-1)*ColNum + j).Interior.Color = ColorIndex(i,j);
    end
end

 

Workbook.Save   % 保存文档