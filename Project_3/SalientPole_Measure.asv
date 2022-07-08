pole = 32;
R_rotor = 5670/2; % mm
ratio = 3/5;

ang1 = [(360/pole)/2]*(pi/180);
ang2 = [(360/pole)/2]*ratio*(pi/180);

x1_x2 = R_rotor*sin(ang1);
h3 = R_rotor*cos(ang1);
Dx = h3/cos(ang2);
x1 = sqrt(Dx^2-h3^2);
x2 = x1_x2 - x1;
h2 = sqrt(R_rotor^2-x1^2)-h3;

HightShoe = R_rotor - h3;
WidthShoe = 2*x1;
