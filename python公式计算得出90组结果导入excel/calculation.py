#导入python模块
from math import *
import xlwt
#建立工作表
workbook = xlwt.Workbook()
#建立表格的第一个sheet
worksheet1 = workbook.add_sheet('first')
#建立表头
worksheet1.write(0,1,'s',)
worksheet1.write(0,2,'τ',)
worksheet1.write(0,3,'θ',)
worksheet1.write(0,4,'x',)
worksheet1.write(0,5,'y',)
worksheet1.write(0,6,'z',)

#全局变量
a = 195.65
b = 49.1025
c = 135.2

for i in range(90):  #添加循环，从0-89，下面的缩进是循环体
    t = 1/90 * (i + 1) #i是从0开始到89，所以加1
    if(t <= 0.125):
        s = 1 / (4 + pi) * (pi * t - 0.25 * sin(radians(4 * pi * t)))
    elif(t >= 0.875):
        s = 1 / (4 + pi) * (pi * t - 0.25 * sin(radians(4 * pi * t)))
    else:
        s = 1/(4+pi)*( 2+pi*t - 9/4*sin( radians((pi + 4*pi*t)/3) ) )
    tao = s * 45
    seita = s * 180
    temp = pi/8 + tao
    x = ( a*cos( radians(temp) ) + b*cos( radians(temp) ) - 180 ) * cos( radians(seita) ) - c*sin( radians(seita) )
    y = a*sin( radians(temp) ) + b*sin( radians(temp) )
    z = ( a*cos( radians(temp) ) + b*cos( radians(temp) ) - 180 ) * cos( radians(seita) ) + c*sin( radians(seita) )
    #下面开始写数据 write(行号，列号，数据) 表格也是从0行0列开始的
    worksheet1.write(i + 1, 0, i+1, )
    worksheet1.write(i + 1, 1, s, )
    worksheet1.write(i + 1, 2, tao, )
    worksheet1.write(i + 1, 3, seita, )
    worksheet1.write(i + 1, 4, x, )
    worksheet1.write(i + 1, 5, y, )
    worksheet1.write(i + 1, 6, z, )
    #下面注释掉了，print函数本来是直接打印的
    # print('\n')
    # print('当T=' + str(t) + '时' +'(第' + str(i+1) + '数)')
    # print('s=' + str(s))
    # print('τ=' + str(tao))
    # print('θ=' + str(seita))
    # print('x=' + str(x))
    # print('y=' + str(y))
    # print('z=' + str(z))
#至此循环体结束

#建立第二个sheet
worksheet2 = workbook.add_sheet('second')
#建立表头
worksheet2.write(0,1,'v',)
worksheet2.write(0,2,'w',)
worksheet2.write(0,3,'β',)

#这个循环和上面的一个一样，做第二个公式
for n in range(90):
    t = 1/90 * (n + 1)
    if(t <= 0.125):
        s = 1 / (4 + pi) * (pi * t - 0.25 * sin(radians(4 * pi * t)))
        v = pi / (4 + pi) * (1 - cos(radians(4 * pi * t)))
    elif(t >= 0.875):
        s = 1 / (4 + pi) * (pi * t - 0.25 * sin(radians(4 * pi * t)))
        v = pi / (4 + pi) * (1 - cos(radians(4 * pi * t)))
    else:
        s = 1 / (4 + pi) * (2 + pi * t - 9 / 4 * sin(radians((pi + 4 * pi * t) / 3)))
        v = pi / (4 + pi) * (1 - 3 * cos(radians((pi + 4 * pi * t) / 3)))
    tao = s * 45
    temp = pi / 8 + tao
    oumige = v*10*pi/3
    beita =atan( c*cos(radians(temp)) /(a*(oumige/(40*pi/3)) + c*sin(radians(temp))) )
    worksheet2.write(n + 1, 0, n + 1, )
    worksheet2.write(n + 1, 1, v, )
    worksheet2.write(n + 1, 2, oumige, )
    worksheet2.write(n + 1, 3, beita, )
    # print('\n')
    # print('当T=' + str(t) + '时' + '(第' + str(n + 1) + '数)')
    # print('v=' + str(v))
    # print('ω=' + str(oumige))
    # print('β=' + str(beita))
#第二个循环结束

#保存成data.xls文件，保存在和程序同一目录下
workbook.save('data.xls')




