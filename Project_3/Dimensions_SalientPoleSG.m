% ratings
p = 32;
pp = p/2;
fe = 50; % Hz
S = 44e6; % Va
pf = 0.9;
Vline =13.8e3; % Vrms
Iline = S/(sqrt(3)*Vline); % Arms

% Armature Current
Jar = 3; % A/mm2
Acu_ar = (Iline/Jar)*1e-6; % mm2

% loadings
A = 65e3; % A/m (rms)
B = 1.05; % Air-gap flux density (Tesla) (peak)
kw1_pseudo = 0.94;

fsyn = fe/pp;
C = (pi^2/sqrt(2))*kw1_pseudo*A*B;
RotorVolume = S/(C*fsyn);
AspectRatio = pi*sqrt(pp)/(4*pp);
% main dimensions
D_bore = (RotorVolume/AspectRatio)^(1/3);
l_core = D_bore*AspectRatio;
Do = 1.10*D_bore;
polePitch = pi*D_bore/p;
% air gap
c_airgap = 4e-7;
minAirgap = c_airgap*polePitch*A/B;
airgap = 16e-3;
% Flux per pole
fluxPole = B*(2/pi)*polePitch*l_core; % Weber

% Turns number
Nph = (Vline/sqrt(3))/(kw1_pseudo*fluxPole*fe*(2*pi/sqrt(2))); % turns
Nph = round(Nph);

% Slot and tooth dimensions
Nturnslot = 1;
Nslot = (Nph*Nturnslot*3*2)/2;
Ku_ar = 0.45;
Aslot_ar = 2*Acu_ar/Ku_ar;
slotPitch = pi*D_bore/Nslot;
slotRatio = 0.4;
toothRatio = 1 - slotRatio;
toothWidth = toothRatio*slotPitch;
slotWidth = slotRatio*slotPitch;
slotHeight = 1.05*Aslot_ar/slotWidth;
BtoothPeak = (pi/2)*fluxPole/(polePitch*toothRatio*l_core);
backcoreHeight = (Do-D_bore-2*slotHeight)/2;
BbackcorePeak = (fluxPole/2)/(backcoreHeight*l_core)*(pi/2);

% MMF
u0 = 4*pi*1e-7;
poleRatio = 2/3;
reluctanceAirgap = airgap/(u0*polePitch*l_core*poleRatio);
airgapMMF = fluxPole*reluctanceAirgap*1.30;
armatureMMF = 2.7*Iline*Nph*kw1_pseudo/p;
cospf = 0.9;
sinpf = sqrt(1-cospf^2);
totalMMF = sqrt(((airgapMMF+armatureMMF*sinpf)^2) + (armatureMMF*cospf)^2);

% Field winding
Jfield = 3;
Ku_field = 0.5;
Nf = 10;
If = totalMMF/Nf;
A_cu_field = Nf*(If/Jfield)*1e-6;
a_field = (poleRatio*polePitch*0.4)/2;
b_field = (A_cu_field/Ku_field)/a_field;

%rmxprt rotor pole info
widthShoe = polePitch*poleRatio;
widthBody = widthShoe - 2*a_field;
heightBody = b_field;
heightShoe = widthShoe*0.15;

