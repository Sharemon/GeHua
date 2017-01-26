function [bz,az]=band_pass_filter(f1,f3,fsl,fsh,rp,rs,Fs)

% f1=7.2e6;f3=10e6;%通带截止频率上下限
% fsl=6.2e6;fsh=11e6;%阻带截止频率上下限
% rp=3;rs=30;%通带边衰减DB值和阻带边衰减DB值
% Fs=100e6;%采样率
%band pass filter
wp1=2*pi*f1/Fs;
wp3=2*pi*f3/Fs;
wsl=2*pi*fsl/Fs;
wsh=2*pi*fsh/Fs;
wp=[wp1 wp3];
ws=[wsl wsh];
%
% 设计切比雪夫滤波器；
[n,wn]=buttord(ws/pi,wp/pi,rp,rs);
[bz,az]=butter(n,wn);
%查看设计滤波器的曲线
[h,w]=freqz(bz,az,256,Fs);
h=20*log10(abs(h));
% figure;plot(w,h);title('所设计滤波器的通带曲线');grid on;
%y=filter(bz1,az1,x);

end