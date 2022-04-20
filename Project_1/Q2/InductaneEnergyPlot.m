%% Arrays and curve fit prepration
SamplePoints =  [0 15 30 45 60 75 80 85 88 90];
L_linear = 1e-3.*[27.194 24.725 21.190 17.537 13.809 10.042 9.233 8.932 8.868 8.854];

a = size(SamplePoints);
a = a(2);
for i = 1:a-1
    SamplePoints(a+i) = 90+SamplePoints(a)-SamplePoints(a-i);
    L_linear(a+i) = L_linear(a-i);
end

a = size(SamplePoints);
a = a(2);
for i = 1:a-1
    SamplePoints(a+i) = 180+SamplePoints(a)-SamplePoints(a-i);
    L_linear(a+i) = L_linear(a-i);
end

a = size(SamplePoints);
a = a(2);
for i = 1:a-1
    SamplePoints(a+i) = 360+SamplePoints(a)-SamplePoints(a-i);
    L_linear(a+i) = L_linear(a-i);
end
SamplePoints = SamplePoints.*(pi/180);
%% Pchip interpolation
SamplePoints_Linspace = linspace(SamplePoints(1),SamplePoints(numel(SamplePoints)),2e2);

L_linear_Interpolated = pchip(SamplePoints,L_linear,SamplePoints_Linspace);
%plot(SamplePoint_Linspace,L_linear_Interpolated);
[L_linear_Derivative,SamplePoints_Derivative] = NumericDerivative(L_linear_Interpolated,SamplePoints_Linspace);
Torque = L_linear_Derivative.*0.5.*(3^2);

%% Plot latex settings
% Some figure formatting
set(groot,'defaulttextinterpreter','latex');  
set(groot, 'defaultAxesTickLabelInterpreter','latex');  
set(groot, 'defaultLegendInterpreter','latex');
figure(1)
subplot(2,1,1)
plot(SamplePoints_Linspace,L_linear_Interpolated*1e3);
legend({'Inductance - Linear core - Interpolation'},'FontSize',18,'Location','northeast')
title('Inductance Variation for Linear Core Material - FEA')
xlabel('Angle $Rad$')
ylabel('Inductance $mH$')
xlim([0 2*pi])
set(findall(gcf,'Type','line'),'LineWidth',5)
set(findall(gcf,'-property','FontSize'),'FontSize',24);

subplot(2,1,2)
plot(SamplePoints_Derivative,Torque*1e3);
legend({'Torque - Linear core - Interpolation'},'FontSize',18,'Location','northeast')
title('Torque Variation w.r.t Angle of Rotor for Linear Core Material - FEA')
xlabel('Angle $Rad$')
ylabel('Torque $mN-m$')
xlim([0 2*pi])
set(findall(gcf,'Type','line'),'LineWidth',5)
set(findall(gcf,'-property','FontSize'),'FontSize',24);

figure(2)
subplot(2,1,1)
plot(SamplePoints_Linspace,L_CurveFit*1e3);
legend({'Inductance - Linear core - Curvefit'},'FontSize',18,'Location','northeast')
title('Inductance Variation for Linear Core Material - FEA')
xlabel('Angle $Rad$')
ylabel('Inductance $mH$')
xlim([0 2*pi])
set(findall(gcf,'Type','line'),'LineWidth',5)
set(findall(gcf,'-property','FontSize'),'FontSize',24);
subplot(2,1,2)
plot(SamplePoints_Linspace,Torque_CurveFit*1e3);
legend({'Torque - Linear core - Curvefit'},'FontSize',18,'Location','northeast')
title('Torque Variation w.r.t Angle of Rotor for Linear Core Material - FEA')
xlabel('Angle $Rad$')
ylabel('Torque $mN-m$')
xlim([0 2*pi])
set(findall(gcf,'Type','line'),'LineWidth',5)
set(findall(gcf,'-property','FontSize'),'FontSize',24);
%% Another Diff
%K = 0.5*9*diff(L_linear_Interpolated)./(SamplePoints_Linspace(2)-SamplePoints_Linspace(1));
K = 0.5*9*diff(L_linear_Interpolated)./diff(SamplePoints_Linspace);
plot(SamplePoints_Linspace(1:(1e7-1)),K);
%% CurveFit - Rad SamplePoint
a = [0.02751 0.008896 0.01176 0.002651 0.001484 0.000292 -0.0002846 0.0005696];
b = [0.2381 1.965 0.4803 0.9596 1.433 7.984 2.548 5.997];
c = [0.07462 1.789 1.694 1.825 1.993 1.673 1.271 1.588];
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