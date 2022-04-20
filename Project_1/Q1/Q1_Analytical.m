%% Sinusoidal Torque Approximation
Lmin = 4.988e-3;
Lmax = 26.528e-3;
Angle = linspace(0,2*pi,100);
L = [(Lmin+Lmax)/2] +[(Lmax-Lmin)/2]*cos(2.*Angle);
% Some figure formatting
set(groot,'defaulttextinterpreter','latex');  
set(groot, 'defaultAxesTickLabelInterpreter','latex');  
set(groot, 'defaultLegendInterpreter','latex');

plot(Angle.*(180/pi),L*1e3);
legend({'Inductance'},'FontSize',14)
title('Inductance Varience Approximation')
xlabel('Angle $(Degree)$')
ylabel('Inductance $(mH)$')
xlim([0 360])
ylim([0 30])
set(findall(gcf,'Type','line'),'LineWidth',4)
set(findall(gcf,'-property','FontSize'),'FontSize',24);

%% Torque
I = 3;
Torque = -(Lmax-Lmin)*sin(2.*Angle)*(I^2)*0.5;
subplot(2,1,1)
plot(Angle.*(180/pi),L*1e3);
legend({'Inductance'},'FontSize',14)
title('Inductance Varience Approximation')
xlabel('Angle $(Degree)$')
ylabel('Inductance $(mH)$')
xlim([0 360])
ylim([0 30])
set(findall(gcf,'Type','line'),'LineWidth',4)

subplot(2,1,2)
plot(Angle.*(180/pi),Torque*1e3);
legend({'Torque'},'FontSize',14)
title('Torque Varience Approximation')
xlabel('Angle $(Degree)$')
ylabel('Torque $(mN-m)$')
xlim([0 360])
ylim([-120 120])
set(findall(gcf,'Type','line'),'LineWidth',4)
set(findall(gcf,'-property','FontSize'),'FontSize',24);