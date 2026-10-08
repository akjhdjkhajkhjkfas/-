function [book] = HolographicCode(k0,kt,angleX,angleY,dx,dy,numberX,numberY,angleLimit,nn);


% % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % 
% % 该函数为基于相位筛选的全息编码计算，针对二维阵列计算
% % 输入变量包括：自由空间波数k0        导行波波数kt
% %              垂直扫描角angleX      水平扫描角angleY
% %              垂直距离dx            水平距离dy
% %              垂直单元数numberX     水平单元数numberY
% %              角度容差angleLimit
% %              极化判断nn，默认应为1，针对三星项目的极化2应改变判断数值为2
% % 输出变量为全息分布码本book
% % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % % %             

u = sin(angleX/180*pi);
v = sin(angleY/180*pi);
pt = uv2phitheta([u;v]);                %归一指向至uv坐标系
phi0 = pt(1)/180*pi;
theta0 = pt(2)/180*pi;                  %单位制为弧度
dxin = -dx*(numberX-1)/2;               %x方向起始距离
dyin = 0;                               %y方向起始距离
xn = [0:1:numberX-1]*dx+dxin;           %x坐标
ym = [0:1:numberY-1]*dy+dyin;           %y坐标,默认阵列基准坐标在上下最中，左右最左
[Xn,Ym] = meshgrid(xn,ym);
Eobj = rad2deg(angle(exp(sqrt(-1)*-k0*(Xn*sin(theta0)*cos(phi0)+Ym*sin(theta0)*sin(phi0)))));
                                        %期望波束场计算
Eref = rad2deg(angle(exp(sqrt(-1)*kt*(Ym)*((-1)^nn))));
                                        %导行调制波场计算
book = abs(Eobj-Eref);                  %两种场角度做差计算分布
book(book<=angleLimit) = 1;
book(book>=360-angleLimit) = 1;
                                        %吻合在容差角度内判定为1
book(book>angleLimit) = 0;              %吻合在容差角度外判定为0
book = book';
