function [bz,az]=high_pass_filter(fp,fs,rp,rs,Fs)

%low pass filter
wp=2*pi*fp/Fs;
ws=2*pi*fs/Fs;
%
% ����б�ѩ���˲�����
[n,wn]=buttord(wp/pi,ws/pi,rp,rs);
[bz,az]=butter(n,wn,'high');
%�鿴����˲���������
[h,w]=freqz(bz,az,256,Fs);
h=20*log10(abs(h));
% figure;plot(w,h);title('������˲����ĵʹ�����');grid on;
%y=filter(bz1,az1,x);

end