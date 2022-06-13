slotPhase = 360*10/21;
n = 5; % Harmonic Number
for k = 1:21
    slot(k) = slotPhase*(k-1)*n;
    slot(k) = mod(slot(k),360);
    %slot(k) = round(slot(k),4,'significant')
    slot(k) = slot(k)*2*pi/360;
end

A = 1*exp(slot(1)*i)-2*exp(slot(2)*i)+2*exp(slot(3)*i)-2*exp(slot(4)*i)+...
    2*exp(slot(5)*i)-2*exp(slot(6)*i)+2*exp(slot(7)*i)-1*exp(slot(8)*i);
A_Magnitude = abs(A);
A_kd = A_Magnitude/14;

%% Kp
A_kp = sin(n*(slotPhase*pi/180)/2);
%%
A_kw = A_kd*A_kp;