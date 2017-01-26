function [bz,az]=low_pass_filter(fp,fs,rp,rs,Fs)

%low pass filter
wp=2*pi*fp/Fs;
ws=2*pi*fs/Fs;
%
% ����б�ѩ���˲�����
[n,wn]=buttord(ws/pi,wp/pi,rp,rs);
[bz,az]=butter(n,wn);
%�鿴����˲���������
[h,w]=freqz(bz,az,256,Fs);
h=20*log10(abs(h));
% figure;plot(w,h);title('������˲����ĵʹ�����');grid on;
%y=filter(bz1,az1,x);

end