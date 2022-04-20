%% Arrays and curve fit prepration
SamplePoints =  [0 15 30 45 60 75 80 85 88 90];
L_nonlinear = 1e-3.*[27.025 24.636 21.103 17.416 13.685 9.900 9.123 8.826 8.762 8.748];

a = size(SamplePoints);
a = a(2);
for i = 1:a-1
    SamplePoints(a+i) = 90+SamplePoints(a)-SamplePoints(a-i);
    L_nonlinear(a+i) = L_nonlinear(a-i);
end

a = size(SamplePoints);
a = a(2);
for i = 1:a-1
    SamplePoints(a+i) = 180+SamplePoints(a)-SamplePoints(a-i);
    L_nonlinear(a+i) = L_nonlinear(a-i);
end

a = size(SamplePoints);
a = a(2);
for i = 1:a-1
    SamplePoints(a+i) = 360+SamplePoints(a)-SamplePoints(a-i);
    L_nonlinear(a+i) = L_nonlinear(a-i);
end
SamplePoints = SamplePoints.*(pi/180);
%% Pchip interpolation
SamplePoints_Linspace = linspace(SamplePoints(1),SamplePoints(numel(SamplePoints)),2e2);

L_nonlinear_Interpolated = pchip(SamplePoints,L_nonlinear,SamplePoints_Linspace);
%plot(SamplePoint_Linspace,L_linear_Interpolated);
[L_linear_Derivative,SamplePoints_Derivative] = NumericDerivative(L_nonlinear_Interpolated,SamplePoints_Linspace);
Torque = L_linear_Derivative.*0.5.*(3^2);

%% Plot latex settings
% Some figure formatting
set(groot,'defaulttextinterpreter','latex');  
set(groot, 'defaultAxesTickLabelInterpreter','latex');  
set(groot, 'defaultLegendInterpreter','latex');
figure(1)
subplot(2,1,1)
plot(SamplePoints_Linspace,L_nonlinear_Interpolated*1e3);
legend({'Inductance - Non-linear core - Interpolation'},'FontSize',18,'Location','northeast')
title('Inductance Variation for Non-linear Core Material - FEA')
xlabel('Angle $Rad$')
ylabel('Inductance $mH$')
xlim([0 2*pi])
set(findall(gcf,'Type','line'),'LineWidth',5)
set(findall(gcf,'-property','FontSize'),'FontSize',24);

subplot(2,1,2)
plot(SamplePoints_Derivative,Torque*1e3);
legend({'Torque - Non-linear core - Interpolation'},'FontSize',18,'Location','northeast')
title('Torque Variation w.r.t Angle of Rotor for Non-linear Core Material - FEA')
xlabel('Angle $Rad$')
ylabel('Torque $mN-m$')
xlim([0 2*pi])
set(findall(gcf,'Type','line'),'LineWidth',5)
set(findall(gcf,'-property','FontSize'),'FontSize',24);

figure(2)
subplot(2,1,1)
plot(SamplePoints_Linspace,L_CurveFit*1e3);
legend({'Inductance - Non-linear core - Curvefit'},'FontSize',18,'Location','northeast')
title('Inductance Variation for Non-linear Core Material - FEA')
xlabel('Angle $Rad$')
ylabel('Inductance $mH$')
xlim([0 2*pi])
set(findall(gcf,'Type','line'),'LineWidth',5)
set(findall(gcf,'-property','FontSize'),'FontSize',24);
subplot(2,1,2)
plot(SamplePoints_Linspace,Torque_CurveFit*1e3);
legend({'Torque - Non-linear core - Curvefit'},'FontSize',18,'Location','northeast')
title('Torque Variation w.r.t Angle of Rotor for Non-linear Core Material - FEA')
xlabel('Angle $Rad$')
ylabel('Torque $mN-m$')
xlim([0 2*pi])
set(findall(gcf,'Type','line'),'LineWidth',5)
set(findall(gcf,'-property','FontSize'),'FontSize',24);
%% Another Diff
%K = 0.5*9*diff(L_linear_Interpolated)./(SamplePoints_Linspace(2)-SamplePoints_Linspace(1));
K = 0.5*9*diff(L_nonlinear_Interpolated)./diff(SamplePoints_Linspace);
plot(SamplePoints_Linspace(1:(1e7-1)),K);
%% CurveFit - Rad SamplePoint
a = [0.02734 0.008905 0.01169 0.002634 0.001475 0.0002811 -0.0002793 0.0005448];
b = [0.2384 1.966 0.4808 0.9606 1.434 7.986 2.546 5.998];
c = [0.07314 1.787 1.691 1.819 1.986 1.661 1.281 1.582];
syms L(x) Ld(x)
L(x) = a(1)*sin(b(1)*x+c(1)) + a(2)*sin(b(2)*x+c(2)) + a(3)*sin(b(3)*x+c(3)) + a(4)*sin(b(4)*x+c(4)) + ...
    a(5)*sin(b(5)*x+c(5)) + a(6)*sin(b(6)*x+c(6)) + a(7)*sin(b(7)*x+c(7)) + a(8)*sin(b(8)*x+c(8));
Ld(x) = diff(L,x);

for i = 1:numel(SamplePoints_Linspace)
    L_Subsvalue(i) = subs(L,x,SamplePoints_Linspace(i));
    Ld_Subsvalue(i) = subs(Ld,x,SamplePoints_Linspace(i));
    Torque_CurveFit(i)  = double(Ld_Subsvalue(i))*0.5*9;
    L_CurveFit(i) = double(L_Subsvalue(i));
end

%% Functions
function [Derivative,DerivativeParameter] = NumericDerivative(Signal,Parameter)
    for i = 1:numel(Parameter)-1
        Derivative(i) = (Signal(i+1)-Signal(i))/(Parameter(i+1)-Parameter(i));
    end
    DerivativeParameter = Parameter(1:(numel(Parameter)-1));
end