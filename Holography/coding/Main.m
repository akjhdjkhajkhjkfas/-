clc;clear all;
close all;

%% 说明
% % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % 
% % 该程序用于计算双极化的全息编码分布码本
% % 极化定义   
% %     极化1指远离三星标识的端口馈电的阵列极化
% %     极化2指靠近三星标识的端口馈电的阵列极化
% % 扫描角正方向定义
% %     水平正方向指极化1端口指向极化2端口的矢量方向
% %     垂直正方向为三星标识首字母S指向尾字母G的矢量方向
% %     angleX1为极化1垂直扫描方向     angleY1为极化1水平扫描方向
% %     angleX2为极化2垂直扫描方向     angleY2为极化2水平扫描方向
% % 默认码本为两极化码本混合，与阵面分布一一对应，能够直接用于上位机
% %         存储位置为"\CodeBook"
% % 测量单一极化需要将另一个极化全关
% % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % 

%% 数值
f = 28;                                 %频率
lamda  = 300/f;
k0 = 2*pi/lamda;                        %自由空间波数
lamdag = 6.673;
kt = 2*pi/lamdag;                       %传输波矢
angleX1 = -5:5:5;
angleY1 = -10:10:10;                          %正方向扫描角度，x为16列阵列合成方向，y为单列阵元组阵方向
angleX2 = -5:5:5;
angleY2 = -10:10:10;                          %正方向扫描角度
numberX = 40;                           %x方向的阵列数
numberY = 64;                           %y方向的单元数
dx = 7;                                 %x方向间距
dy = 2;                                 %y方向间距
angleLimit = 90;                        %容差角度，±90

%% 计算码本
for i = 1:length(angleX1)
    for j = 1:length(angleY1)
        for k = 1:length(angleX2)
            for h = 1:length(angleY2)
                nn = 1;                 %极化1码本
                [book1] = HolographicCode(k0,kt,angleX1(i),angleY1(j),dx,dy,numberX,numberY,angleLimit,nn);
%                 book1 = zeros(height(book1),length(book1));         %极化1全关
                nn = 2;                 %极化2码本
                [book2] = HolographicCode(k0,kt,angleX2(k),angleY2(h),dx,dy,numberX,numberY,angleLimit,nn);
%                 book2 = zeros(height(book2),length(book2));         %极化2全关                                    
                book = [];
                for u = 1:numberX
                    book(u*2-1,:) = book1(u,:);
                    book(u*2,:) = book2(u,:);
                end
                
                name = strcat('\CodeBook\',num2str(angleX1(i)),'_',num2str(angleY1(j)),'@',num2str(angleX2(k)),'_',num2str(angleY2(h)),'.xls');
                xlswrite(name,book');
            end
        end
    end
end







