clear all
clc
close all

lp_num = 18;
right = 0;

coefs = zeros(lp_num,2);

alpha = rand(1,1)*2*pi;

for dif = 0:20:1000
% for phase = 0:360
for loop = 1:lp_num

%parameter
N = 125000;
f_carry1 = 10e6+dif/2;
f_carry2 = 10e6-dif/2;
fs_virtual = 180e6;
fs_real_std = 7.2e6;
t = 1:N;

% dec_rate = fs_virtual / fs_real_std;
% dec_rate = floor(dec_rate);
% fs_real = fs_virtual/dec_rate;
% f_carry = 2*fs_real;
fs_real = fs_real_std;
dec_rate = fs_virtual / fs_real_std;

%%
%noise in 7.2MHz - 12MHz
% % load noise
% load noise.mat

% create noise with a filter
noise = randn(1,2*N);
% noise2 = randn(1,N);
[b,a] = band_pass_filter(8.5e6, 10.5e6, 6e6, 12e6, 3, 40, fs_virtual);
noise = filter(b,a,noise);
% save 'noise.mat' noise
% figure
% plot(noise)
% NOISE = abs(fft(noise));
% figure
% plot((t-1)/N*fs_virtual, NOISE);


%%
%carry wave
cw1 = sin(2*pi*f_carry1/fs_virtual*(t-1));
cw2 = sin(2*pi*f_carry2/fs_virtual*(t-1) + 0/360*2*pi + alpha);
cw22 = sin(2*pi*f_carry2/fs_virtual*(t-1) + 90/360*2*pi + alpha);
% figure()
% plot(cw1(1:100))
% hold on
% plot(cw2(1:100),'r')

%%
% modulation
x1 = cw1.* (noise(1:N));
% pos = floor(rand(1,1)*(N*9/10));
pos=loop;
x2 = cw2.* noise(pos+1:(N + pos));
x22 = cw22.* noise(pos+1:(N + pos));
% figure()
% plot(x1)
% hold on
% plot((x2),'r')
% 
% X1=abs(fft(x1));
% figure()
% plot((t-1)/N*fs_virtual,X1)

%%
% 3MHz LPF
[b,a] = low_pass_filter(1.9e6,3.3e6,3,40,fs_virtual);
x_lpf1 = filter(b,a,x1);
x_lpf2 = filter(b,a,x2);
x_lpf22 = filter(b,a,x22);
% [b,a] = high_pass_filter(2e6,0.1e6,3,40,fs_virtual);
% x_lpf1 = filter(b,a,x_lpf1);
% x_lpf2 = filter(b,a,x_lpf2);
% x_lpf22 = filter(b,a,x_lpf22);
% x_lpf1 = x1;
% x_lpf2 = x2;
% 
% figure()
% plot(x_lpf1)
% hold on
% plot((x_lpf2),'r')
% 
% figure()
% X_LPF1 = fft(x_lpf1);
% plot(abs(X_LPF1))

for i = 1: N/dec_rate
    x_resample_lpf1(i) = x_lpf1((i-1)*dec_rate+1);
    x_resample_lpf2(i) = x_lpf2((i-1)*dec_rate+1);
    x_resample_lpf22(i) = x_lpf22((i-1)*dec_rate+1);
end

% x_resample_lpf1 = x_lpf1;
% x_resample_lpf2 = x_lpf2;

% figure()
% plot(x_resample_lpf1)
% hold on
% plot((x_resample_lpf2),'r')

% X_RS_LPF1 = fft(x_resample_lpf1);
% figure()
% plot(abs(X_RS_LPF1))
% 

% x_resample_lpf1 = abs(x_resample_lpf1);
% x_resample_lpf2 = abs(x_resample_lpf2);
% x_resample_lpf22 = abs(x_resample_lpf22);

% self = xcorr(x_resample_lpf1);
each = xcorr(x_resample_lpf1,x_resample_lpf2);
each2 = xcorr(x_resample_lpf1,x_resample_lpf22);

% figure()
% plot(self)
% figure()
% plot(abs(each))
[c,i] = max(abs(each));
i;
[c2,i2] = max(abs(each2));
i2;
% abs(i-10000)
% pos/20
% coef = corrcoef(x_resample_lpf1(2:end),x_resample_lpf2(1:end-1))
% coefs(loop) = coef(1,2);
% 
if (1)%(floor(pos/20) <= abs(i-10000) && ceil(pos/20) >= abs(i-10000))
    right=right+1;
    mid2 = x_resample_lpf2(1:end-abs(i-10000 * N/250000));
    mid22 = x_resample_lpf22(1:end-abs(i2-10000 * N/250000));
    mid1 = x_resample_lpf1(abs(i-10000 * N/250000)+1:end);
    mid12 = x_resample_lpf1(abs(i2-10000 * N/250000)+1:end);
%     figure
%     plot(mid1)
%     hold on
%     plot(mid2,'r')
    coef = corrcoef(mid1,mid2);
    coef2 = corrcoef(mid12,mid22);
    ok(loop) = max(abs(coef(1,2)),abs(coef2(1,2)));
    coefs(loop,:) = [coef(1,2),coef2(1,2)];
    poss(loop) = pos;
    is(loop) = i;
end



end
% 
% figure
% plot(ok)
% figure
% plot((1:lp_num)/18*360,coefs,'*-')
% grid on
% legend('0°载波','90°载波')

% % 
% plot((0:360)/200*360,abs(coefs))
% title('载波相位偏移 vs 相关系数')
% xlabel('角度 0 - 360°')
% ylabel('相关系数')
% grid on
% right

maxx(dif/20+1) = max(ok);

end

figure()
plot((0:20:1001),maxx)
grid on
%%
%analysis
% X1=fft(x1);
% X2=fft(x2);
% 
% figure
% plot(abs(X1))
% hold on
% plot(abs(X2),'r')
% 
% CW1=fft(cw1);
% CW2=fft(cw2);
% 
% figure
% plot(angle(CW1))
% hold on
% plot(angle(CW2),'r')
% 
% figure
% plot(abs(CW1))
% hold on
% plot(abs(CW2),'r')
% X_RS_LPF1 = fft(x_resample_lpf1);
% figure()
% plot(abs(X_RS_LPF1))