% Semiautomated fitting of the single-shell CM model to data collected by
% the DEPtech 3DEP electrophysiology reader. Modified from the adamstn2
% script (github.com/adamstn2/Core-shell-spherical-DEP-polarization-model, 
% https://doi.org/10.1002/elps.202300202)
% to avoid the csv input and adopt the MPH-group variable conventions.
%
% Author: MP Johnson
% Release: 2025
% https://github.com/mpjdep/3DEP
%
% First publication to utilise this code: http://dx.doi.org/10.1038/s41598-025-06371-5

%% Constants 
t = 7e-9;               % Cell membrane thickness (m). Default: 7e-9
e0 = 8.85e-12;          % Permittivity of free space, (F/m). Universal Constant.
e3 = 78*e0;             % Medium permittivity, (F/m). Default 78*e0 for aquaous solutions.

%% ---------------------- EDITABLE VALUES ----------------------------------
r = (8.2810e-6)/1;      % Cell radius (m)
q3 = 0.1609;            % Medium conductivity, (S/m), (100uS/cm = 10mS/m)

% Paste data from 3DEP output below.
data = [	
10000	-1.153157232
15569.48842	-1.396554416
37741.83646	-2.012785069
58762.10858	-1.566495879
91489.59691	-1.405231661
142444.622	-1.171034524
221778.9892	-1.269875506
537612.1627	-1.001148874
837034.6342	0.032898572
1303220.105	0.046753299
% 2029047.033	-0.107002006
3159122.428	0.263902452
4918592.006	0.68859104
7657996.129	1.153263975
11923108.2	0.489371582
18563669.51	0.866357011
28902683.75	0.314277773
45000000	0.323348671
];

x0    = [50*e0,    15*e0,     q3,    1e-5,     5];      % INITIAL GUESSES
x_min = [10*e0,    0*e0,     0.001,   1e-14,      3];   % MINIMUM BOUNDS
x_max = [160*e0,    200*e0,    1.6,    1e-2,      20];  % MAXIMUM BOUNDS
% -----------------------------------------------------------------------

%% The Processing - No editing below
x1 = data(1:end,1);
y1 = data(1:end,2);
options = optimset('PlotFcns',@optimplotfval);
fun     = @(x)residual(x1,y1,e0,e3,q3,r,t,x); % x = [e1,e2,q1,q2]
xfinal  = fminsearchbnd(fun, x0, x_min, x_max, options);
e1 = xfinal(1);             % Cytoplasmic permittivity, (F/m)
e2 = xfinal(2);             % Membrane permittivity, (F/m)
q1 = xfinal(3);             % Cytoplasmic conductivity, (S/m)
q2 = xfinal(4);             % Membrane conductivity, (S/m)
y1 = y1./xfinal(5);
xfinal_ymodel   = depf(e0,e3,e1,e2,q1,q2,q3,r,t,x1);
xfinal_SSE      = sum((y1-xfinal_ymodel).^2);
xfinal_SSTO     = sum((y1-mean(y1)).^2); 
xfinal_r_square = 1-xfinal_SSE/xfinal_SSTO; 
disp(['e1 = ',  num2str(xfinal(1)/e0), ' *e0']);
disp(['e2 = ',  num2str(xfinal(2)/e0), ' *e0']);
disp(['q1 = ',  num2str(xfinal(3)), ' S/m']);
disp(['q2 = ',  num2str(xfinal(4)), ' S/m']);
disp(['factor = ',  num2str(1/xfinal(5))]);
disp(['r2 = ',  num2str(xfinal_r_square)]);
disp(' ');
Cm = xfinal(2)/t;
Gm = xfinal(4)/t;

%% Figure
freq = 10.^linspace(log10(1e3),log10(1e10),10000); 
RealFcm1 = depf(e0,e3,e1,e2,q1,q2,q3,r,t,freq);
figure(1)
semilogx(x1,y1,'o', freq,(RealFcm1),'r')
xlim([1e3, 1e8]); ylim([-0.5, 1.5]);
title ('DEP Spectrum');
xlabel('Frequency (Hz)');
ylabel('Re(F_C_M)');

%% Transient Slope Calculation. E Egun and TNG Adams: https://doi.org/10.3390/biophysica4040045
slopecalc_y = RealFcm1(1:find(RealFcm1 == max(RealFcm1),1));
slopecalc_x = freq(1:find(RealFcm1 == max(RealFcm1),1));
slopecalc_lb = min(slopecalc_y) + 0.2 * (max(slopecalc_y)-min(slopecalc_y));
slopecalc_ub = min(slopecalc_y) + 0.8 * (max(slopecalc_y)-min(slopecalc_y));
slopecalc_x = slopecalc_x(slopecalc_y > slopecalc_lb & slopecalc_y < slopecalc_ub);
slopecalc_y = slopecalc_y(slopecalc_y > slopecalc_lb & slopecalc_y < slopecalc_ub);
slopecalc_x = log10(slopecalc_x);
[P,S] = polyfit(slopecalc_x,slopecalc_y,1);
r2_TransSlope = 1 - (S.normr/norm(slopecalc_y - mean(slopecalc_y)))^2;

%% Functions
% Author: John D'Errico
% E-mail: woodchips@rochester.rr.com
% Release: 4
% Release date: 7/23/06
% License @ https://uk.mathworks.com/matlabcentral/fileexchange/8277-fminsearchbnd-fminsearchcon
function [x,fval,exitflag,output] = fminsearchbnd(fun,x0,LB,UB,options,varargin)
    xsize = size(x0);
    x0 = x0(:);
    n=length(x0);
    if (nargin<3) || isempty(LB)
      LB = repmat(-inf,n,1);
    else
      LB = LB(:);
    end
    if (nargin<4) || isempty(UB)
      UB = repmat(inf,n,1);
    else
      UB = UB(:);
    end
    if (n~=length(LB)) || (n~=length(UB))
      error 'x0 is incompatible in size with either LB or UB.'
    end
    if (nargin<5) || isempty(options)
      options = optimset('fminsearch');
    end
    params.args = varargin;
    params.LB = LB;
    params.UB = UB;
    params.fun = fun;
    params.n = n;
    params.xsize = xsize;
    params.OutputFcn = [];
    params.BoundClass = zeros(n,1);
    for i=1:n
      k = isfinite(LB(i)) + 2*isfinite(UB(i));
      params.BoundClass(i) = k;
      if (k==3) && (LB(i)==UB(i))
        params.BoundClass(i) = 4;
      end
    end
    x0u = x0;
    k=1;
    for i = 1:n
      switch params.BoundClass(i)
        case 1
          if x0(i)<=LB(i)
            x0u(k) = 0;
          else
            x0u(k) = sqrt(x0(i) - LB(i));
          end
          k=k+1;
        case 2
          if x0(i)>=UB(i)
            x0u(k) = 0;
          else
            x0u(k) = sqrt(UB(i) - x0(i));
          end
          k=k+1;
        case 3
          if x0(i)<=LB(i)
            x0u(k) = -pi/2;
          elseif x0(i)>=UB(i)
            x0u(k) = pi/2;
          else
            x0u(k) = 2*(x0(i) - LB(i))/(UB(i)-LB(i)) - 1;
            x0u(k) = 2*pi+asin(max(-1,min(1,x0u(k))));
          end
          k=k+1;
        case 0
          x0u(k) = x0(i);
          k=k+1;
        case 4
      end
    end
    if k<=n
      x0u(k:n) = [];
    end
    if isempty(x0u)
      x = xtransform(x0u,params);
      x = reshape(x,xsize);
      fval = feval(params.fun,x,params.args{:});
      exitflag = 0;
      output.iterations = 0;
      output.funcCount = 1;
      output.algorithm = 'fminsearch';
      output.message = 'All variables were held fixed by the applied bounds';
      return
    end
    if ~isempty(options.OutputFcn)
      params.OutputFcn = options.OutputFcn;
      options.OutputFcn = @outfun_wrapper;
    end
    [xu,fval,exitflag,output] = fminsearch(@intrafun,x0u,options,params);
    x = xtransform(xu,params);
    x = reshape(x,xsize);
      function stop = outfun_wrapper(x,varargin);
        xtrans = xtransform(x,params);
        stop = params.OutputFcn(xtrans,varargin{1:(end-1)});
      end
end
function fval = intrafun(x,params)
    xtrans = xtransform(x,params);
    fval = feval(params.fun,reshape(xtrans,params.xsize),params.args{:});
end
function xtrans = xtransform(x,params)
    xtrans = zeros(params.xsize);
    k=1;
    for i = 1:params.n
      switch params.BoundClass(i)
        case 1 % lower bound only
          xtrans(i) = params.LB(i) + x(k).^2;
          k=k+1;
        case 2 % upper bound only
          xtrans(i) = params.UB(i) - x(k).^2;
          k=k+1;
        case 3 % lower and upper bounds
          xtrans(i) = (sin(x(k))+1)/2;
          xtrans(i) = xtrans(i)*(params.UB(i) - params.LB(i)) + params.LB(i);
          xtrans(i) = max(params.LB(i),min(params.UB(i),xtrans(i)));
          k=k+1;
        case 4 % fixed variable, bounds are equal, set it at either bound
          xtrans(i) = params.LB(i);
        case 0 % unconstrained variable.
          xtrans(i) = x(k);
          k=k+1;
      end
    end
end

% Author: Drs. Tayloria N.G. Adams, Tunglin "Anthony" Tsai 
% MIT License
% https://github.com/adamstn2/Core-shell-spherical-DEP-polarization-model
function RealCM1 = depf(~,Emed,Ecyt,Emem,Qcyt,Qmem,Qmed,Rout,t,freq)
    Rin = Rout - t;
    a = Rout/Rin;
    w = (2*pi).*freq;
    % ------------- Core-shell Spherical Model Equations (DO NOT CHANGE) -------------
    ECcyt = Ecyt + (Qcyt./((sqrt(-1)).*w));  % Complex permittivitty of the cytoplasm
    ECmem = Emem + (Qmem./((sqrt(-1)).*w));  % Complex permittivitty of the cell membrane
    ECmed = Emed + (Qmed./((sqrt(-1)).*w));  % Complex permittivitty of the medium, varying medium cond
    ECeff = ECmem.*((a^3+(2.*((ECcyt-ECmem)./(ECcyt+ 2.*ECmem))))./(a^3-((ECcyt-ECmem)./(ECcyt+2.*ECmem))));
    CM1     = (ECeff-ECmed)./(ECeff+(2.*ECmed));
    RealCM1 = real(CM1);
end
function [err, RealFcm1] = residual(x1,y1,Evac,Emed,Qmed,Rout,t,x0)
    RealFcm1 = depf(Evac,Emed,x0(1),x0(2),x0(3),x0(4),Qmed,Rout,t,x1);
    err = x0(5)*sum(abs(RealFcm1-y1./x0(5)));
end


