function [bz,az]=high_pass_filter(fp,fs,rp,rs,Fs)

%low pass filter
wp=2*pi*fp/Fs;
ws=2*pi*fs/Fs;
%
% 设计切比雪夫滤波器；
[n,wn]=buttord(wp/pi,ws/pi,rp,rs);
[bz,az]=butter(n,wn,'high');
%查看设计滤波器的曲线
[h,w]=freqz(bz,az,256,Fs);
h=20*log10(abs(h));
% figure;plot(w,h);title('所设计滤波器的低带曲线');grid on;
%y=filter(bz1,az1,x);

end