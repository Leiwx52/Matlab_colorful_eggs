function photo2excel(filename)
% ����MATLABͨ����ɫ�ķ�ʽ��ͼ����Excel�ĵ�(��תExcelƴͼ)
%   photo2excel(filename)  �������filenameΪ����ͼƬ�ļ�����·�����ַ���
%   photo2excel(I)         �������IΪͼ��Ҷ�ֵ����
%   
%   ע��ͼ������ǻҶ�ͼ������ͼ��
%   Copyright 2010 - 2011 xiezhh. 
%   $Revision: 1.0.0.0 [        DISCUZ_CODE_1        ]nbsp; $Date: 2011/03/12 13:41:00 $

 

% �ж�����������ͣ�����ͼ��Ҷ�ֵ����
if ischar(filename) && exist(filename,'file');
    I = imread(filename);
elseif isnumeric(filename)
    I = filename;
else
    error('����������Ͳ�ƥ��');
end
I = uint8(I);

 

% �ж�ͼ������
[m,n,k] = size(I);
if k == 1
    I = repmat(I,[1 1 3]);
end
if k ~= 1 && k ~= 3
    error('ͼ�����Ͳ�ƥ�䣬ֻ���ǻҶ�ͼ������ͼ��');
end

 

% �趨Ҫ�����Excel�ļ�����·��
filespec_user = [pwd '\�ҵ���Ƭ.xls'];

 

% �ж�Excel�Ƿ��Ѿ��򿪣����Ѵ򿪣����ڴ򿪵�Excel�н��в���������ʹ�Excel
try
    % ��Excel�������Ѿ��򿪣���������Excel
    Excel = actxGetRunningServer('Excel.Application');
catch
    % ����һ��Microsoft Excel�����������ؾ��Excel
    Excel = actxserver('Excel.Application'); 
end;

 

% ����Excel������Ϊ�ɼ�״̬
Excel.Visible = 1;

 

% �½�һ���������������棬�ļ���Ϊ�ҵ���Ƭ.Excel
Workbook = Excel.Workbooks.Add;
%Workbook.SaveAs(filespec_user);

 

% ���ص�ǰ��������
Sheets = Excel.ActiveWorkbook.Sheets;
Sheet1 = Sheets.Item(1);    % ���ص�1�������

 

% �����иߺ��п�
SC1 = Sheet1.Cells;
% �����и�
SC1.RowHeight = 1.5;
% �����п�
SC1.ColumnWidth = 0.15;

 

% ͨ����ɫ�ķ�ʽ��ͼ
ColNum = SC1.Columns.Count;
RowNum = SC1.Rows.Count;
I = double(I);
ColorIndex = I(:,:,1) + 256*I(:,:,2) + 65536*I(:,:,3);
for j = 1:n
    for i = 1:m
        SC1.Item((i-1)*ColNum + j).Interior.Color = ColorIndex(i,j);
    end
end

 

Workbook.Save   % �����ĵ�