import time

import pyvisa as visa


class Chamber:

    def __init__(self, address):
        rm = visa.ResourceManager()
        self._chamber = rm.open_resource(address)

    def setConsMode(self):
        """
        设置定值工作模式
        :return:
        """
        self._chamber.write("MODE, CONSTANT")

    def setTemp(self, temp):
        self._chamber.write(f'TEMP, S{temp} H100.0 L-40.0')
        print("TEMP, S%f H100.0 L-40.0" % temp)
        # print('Set temp %f' % temp)
        time.sleep(0.5)

    def powerOff(self):
        self._chamber.write("POWER, OFF")

    def powerOn(self):
        """
        启动定值运行
        :return:
        """
        self._chamber.write("POWER, ON")
        print('Chamber Power on')
        time.sleep(5)

    def getCurrentT(self):
        """
        查询当前温度，返回结果样式‘测量温度，设置温度，高温报警温度，低温报警温度’
        :return:
        """
        data = self.getTstr()
        print(data)
        temp_str = data.split(",")[0]
        temp = float(temp_str)
        return temp

    def getTstr(self):
        strT = self._chamber.query("TEMP?")
        # 剔除一些异常数据
        if 'OK' in strT or 'O' == strT[0]:
            print('error str ' + strT)
            strT = self.getTstr()
        return strT

    def getAvgTof10S(self):
        """
        获取10秒内的平均温度，用于判断温度是否稳定
        :return:
        """
        cycle = 0
        temp = 0
        while cycle < 10:
            temp += self.getCurrentT()
            cycle += 1
            time.sleep(1)
        return temp/10

