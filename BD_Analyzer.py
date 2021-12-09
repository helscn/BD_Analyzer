
# -*- coding: utf-8 -*-

import struct
import csv
import re
import pyqtgraph as pg
from PyQt5.QtGui import QIcon, QColor
from PyQt5.QtWidgets import QMainWindow, QApplication, QMessageBox, QInputDialog, QFileDialog, QTableWidgetItem
from PyQt5 import QtCore, QtGui, QtWidgets
from ui.mainWindow import Ui_MainWindow
import sys
from os import path

BASE_DIR = path.dirname(__file__)
sys.path.insert(0, path.join(BASE_DIR,'ui'))


def parseNumber(s, length=3, lead_zero=True):
    """
    根据Excellon指令中的文本坐标数据识别转换为实际坐标值
    @param s: 需要识别的文本坐标数据
    @type s: str
    @param length: 数据位数长度，当lead_zero为True时代表整数位数，当lead_zero为False(即trail_zero)时代表小数位数
    @type length: int
    @param lead_zero: 设置数据是否为前导零格式，值为Fasle时表示数据为后补零格式，此时length指定了数据的小数位数
    @type lead_zero: bool
    @return: 根据文本数据识别出的实际数值
    @rtype: float
    """

    if not s:
        return 0

    # 删除文本头尾的空格
    s = re.sub(r'(^\s+|\s+$)', '', s)
    # 如果文本中含有小数点直接转换为浮点数
    if '.' in s:
        return float(s)
    # 判断数据的正负号
    if s[0] == '-':
        sign = -1
        s = s[1:]
    else:
        sign = 1
    if lead_zero:
        # 识别前导零格式数据，其中length为整数位数
        if len(s) < length:
            # 当数据位数小于整数位数时根据缺少的位置进行进位
            v = sign * float(s) * pow(10, length-len(s))
        else:
            # 数据位数大于等于整数位时在整数位后插入小数点
            v = sign * float(s[0:length]+'.'+s[length:])
    else:
        # 将识别后导零格式数据，其中length为小数位数
        v = sign * float(s)/pow(10, length)
    return v


def packPos(x, y):
    # 将X、Y程序坐标pack为字节串
    if type(x) is str:
        x = parseNumber(x, 3, True)
    if type(y) is str:
        y = parseNumber(y, 3, True)
    x = int(round(x*1000))
    y = int(round(y*1000))
    return struct.pack('<ii', x, y)


def prgPos(x, y):
    # 将X,Y坐标转换为程序指令
    if abs(x) >= pow(10, 3) or abs(y) >= pow(10, 3):
        raise ValueError('坐标值过大，无法转换为程序坐标({},{})'.format(x, y))
    signx = '-' if x < 0 else ''
    signy = '-' if y < 0 else ''
    x = abs(int(round(x*1000)))
    y = abs(int(round(y*1000)))
    x = str(x).rjust(6, '0').rstrip('0')
    y = str(y).rjust(6, '0').rstrip('0')
    x = x if x else '0'
    y = y if y else '0'
    return 'X{}{}Y{}{}'.format(signx, x, signy, y)


def loadPrgData(file_path):
    '''读取钻孔程序，并保存每个坐标的刀具信息及刀具对应的M18深度信息'''
    tools = {}
    m18 = {}
    tool_size = {}
    toolPattern = re.compile(r'^T\d{1,2}$')
    posPattern = re.compile(r'^X(-?\d*)Y(-?\d*)$')
    m18Pattern = re.compile(r'^M18Z-?\d*\.?\d*$')
    toolSizePattern = re.compile(r'^(T\d{1,2})C\d*\.?\d+$')
    curTool = None
    with open(file_path, 'r') as f:
        for line in f:
            line = line.strip()
            if line:
                isPos = posPattern.match(line)
                if isPos:
                    x, y = isPos.groups()
                    pos = packPos(x, y)
                    if curTool:
                        tools[pos] = curTool
                else:
                    isToolSize = toolSizePattern.match(line)
                    if isToolSize:
                        tool_size[isToolSize.groups()[0]] = line
                    elif toolPattern.match(line):
                        curTool = line
                    elif m18Pattern.match(line):
                        if curTool:
                            m18[curTool] = line
    return (tool_size, tools, m18)


def loadBackDrillData(files):
    '''读取Schmoll机背钻加工记录数据文件'''
    if type(files) is str:
        files = [files]
    files.sort()
    pattern = re.compile(r'^[-+]?\d+\.?\d*$')
    filePattern = re.compile(r'.*pos(\d\d)(\d\d)\.dat$')

    last_H = None
    base = 0
    result = []
    flags = {}
    for file_path in files:
        QApplication.processEvents()
        isDataFile = filePattern.match(file_path)
        if not path.isfile(file_path):
            raise ValueError('未找到加工记录文件路径:\n'+file_path)
        if isDataFile:
            month, day = isDataFile.groups()
        else:
            raise ValueError('无法通过文件名获取加工日期:\n'+file_path)
        with open(file_path, newline='') as f:
            csvfile = csv.reader(f, delimiter=';')

            # 读取 header
            header = next(csvfile)
            # 加工记录总栏位数为 轴数*3+5
            if len(header) <= 8 or header[0]!="#Date" or header[1]!="N":
                raise ValueError('无法识别的背钻记录文件:\n'+file_path)

            for r in csvfile:
                # 是否为新程序起始孔
                flag = False

                # 加工日期时间
                r[0] = '{}/{} {}'.format(month, day, r[0])

                # 孔编号N
                current_H = int(r[1])
                if last_H and current_H < last_H:
                    if last_H == 65535:
                        base += 65536
                    else:
                        base = 0
                        flag = True
                r[1] = current_H+base

                # 将空数据填充为None
                for i, val in enumerate(r):
                    if (type(val) is str) and pattern.match(val):
                        val = float(val)
                        if i > 4 and val == 0:
                            val = None
                        r[i] = val

                # 将加工记录栏位扩充至6轴
                while len(r)<23:
                    r.append(None)                

                # 如果当前记录与前一个记录是同一个孔，则合并记录文件
                if len(result) > 0:
                    r0 = result[-1]
                    if r0[1] == r[1] and r0[3] == r[3] and r0[4] == r[4]:
                        for c in range(5, 23):
                            r0[c] = r[c] or r0[c]
                        continue
                result.append(r)
                last_H = current_H
                if not flags or flag == True:
                    flags[r[0]] = len(result)-1
    return result, flags


def calc_outliers(data, lag, tol):
    # 增加计算异常点的权重栏位并保存权重数据
    length = len(data)

    weight = []
    for i in range(length):
        # 6轴加工数据
        weight.append([0]*6)

    for step in range(1, lag+1):
        i = 0
        while i < length-step:
            # 6轴加工数据列位置
            for z in range(6, 23, 3):
                # 计算保存异常点权重的列索引
                col = round((z-6)/3)
                if data[i][z] and data[i+step][z]:
                    if abs(data[i][z]-data[i+step][z]) > tol:
                        weight[i][col] += 1
                        weight[i+step][col] += 1
            i += 1
    return weight


def getIntPos(x, y):
    # 将坐标*1000并转换为int
    if type(x) is str:
        x = float(x)
    if type(y) is str:
        y = float(y)
    return (int(round(x*1000)), int(round(y*1000)))


def judgeShift(data_pos, prg_pos):
    # 判断两组坐标数组是否呈线性一致
    if len(data_pos) != len(prg_pos):
        return False
    shiftX = None
    shiftY = None
    for m, n in zip(data_pos, prg_pos):
        curX = m[0]-n[0]
        curY = m[1]+n[1]
        if shiftX is None or shiftY is None:
            shiftX = curX
            shiftY = curY
        if shiftX != curX or shiftY != curY:
            return None
    return (shiftX/1000, shiftY/1000)


def calc_shift(prg_file, data, sample=10):
    # 根据程序坐标和加工记录坐标自动计算坐标偏移量
    #   shiftX = 机台坐标X - 程序坐标X
    #   shiftY = 机台坐标Y + 程序坐标Y

    prg_pos = []
    curTool = None
    toolPattern = re.compile(r'^T0?[2-9]$')
    posPattern = re.compile(r'^X(-?\d*)Y(-?\d*)$')
    with open(prg_file, 'r') as prg:
        for line in prg:
            if toolPattern.match(line):
                curTool = line
            elif curTool and posPattern.match(line):
                x, y = posPattern.match(line).groups()
                x = parseNumber(x, length=3, lead_zero=True)
                y = parseNumber(y, length=3, lead_zero=True)
                x, y = getIntPos(x, y)
                prg_pos.append((x, y))
            if len(prg_pos) == sample:
                break

    data_pos = []
    for r in data:
        try:
            x, y = getIntPos(float(r[3]), float(r[4]))
            data_pos.append((x, y))
        except Exception:
            continue
        if len(data_pos) > sample:
            data_pos.pop(0)
        shiftXY = judgeShift(data_pos, prg_pos)
        if shiftXY:
            return shiftXY


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setupUi(self)
        # loadUi(path.join(BASE_DIR, "mainWindow.ui"), self)

        self.setWindowTitle("背钻加工记录分析程序")
        # self.setWindowIcon(QIcon(path.join(BASE_DIR, "icon.ico")))
        self.plot = pg.PlotWidget()
        self.graphLayout.addWidget(self.plot)
        self.lastClickedPoints = []
        self.clickedPen = pg.mkPen('g', width=3)
        self.okBrush = pg.mkBrush(255, 255, 255, 180)
        self.ngBrush = pg.mkBrush(255, 0, 0, 255)
        self.markerSize = 8
        self.data = []
        self.result = []
        self.toolPos = {}
        self.toolSize = {}
        self.toolDepth = {}
        self.posShifted = False
        self.init_event()

    def init_event(self):
        self.btnShowPlot.clicked.connect(self.show_plot)
        self.btnLoadPrg.clicked.connect(self.select_prg)
        self.btnLoadData.clicked.connect(self.select_data)
        self.btnOutputOutlier.clicked.connect(self.output_prg)

    def disableButton(self):
        self.btnLoadPrg.setDisabled(True)
        self.btnLoadData.setDisabled(True)
        self.btnShowPlot.setDisabled(True)
        self.btnOutputOutlier.setDisabled(True)

    def enableButton(self):
        self.btnLoadPrg.setEnabled(True)
        self.btnLoadData.setEnabled(True)
        self.btnShowPlot.setEnabled(True)
        self.btnOutputOutlier.setEnabled(True)

    def select_prg(self):
        self.disableButton()
        fileName, filetype = QFileDialog.getOpenFileName(self,
                                                         "选取钻孔程序",
                                                         "./",
                                                         "背钻程序 (*.B00;*.BM0;*.BA0;*.BB0);;所有文件 (*)")
        if fileName and path.isfile(fileName):
            self.toolSize, self.toolPos, self.toolDepth, = loadPrgData(
                fileName)
            if len(self.toolSize) == 0 or len(self.toolPos) == 0 or len(self.toolDepth) == 0:
                QMessageBox.critical(
                    self, '未知数据格式', '无法识别的背钻加工记录文件！', QMessageBox.Ok)
                self.prgFilePath.setText('')
            else:
                self.prgFilePath.setText(fileName)
            self.posShifted = False
            self.result = []
        self.enableButton()

    def select_data(self):
        self.disableButton()
        fileNames, filetype = QFileDialog.getOpenFileNames(self,
                                                           "选取钻孔程序",
                                                           "./",
                                                           "加工记录文件 (*.dat)")
        if fileNames:
            try:
                self.data, flags = loadBackDrillData(fileNames)
                selectedTime, okPressed = QInputDialog.getItem(
                    self, "选择时间", "请选择加工板的起始加工时间以筛选加工数据:", flags.keys(), 0, False)
                if okPressed:
                    start = flags[selectedTime]
                    end = len(self.data)
                    for key, index in flags.items():
                        if start < index < end:
                            end = index
                    self.data = self.data[start:end]
                    self.dataFilePath.setText(','.join(fileNames))
                else:
                    self.data = []
                    self.dataFilePath.setText('')
            except Exception as e:
                QMessageBox.critical(
                    self, '无法识别背钻数据文件', e.args[0], QMessageBox.Ok)
                self.data = []
                self.dataFilePath.setText('')

            self.posShifted = False
            self.result = []
        self.enableButton()

    def output_prg(self):
        if not self.result:
            QMessageBox.information(
                self, '提示', '请先计算分析测高异常孔.', QMessageBox.Yes)
            return
        selected = QFileDialog.getSaveFileName(self, '输出返工加工程序')
        if selected:
            try:
                with open(selected[0], 'w', newline='\n') as f:
                    toolPattern = re.compile(r'^T\d{1,2}$')
                    with open(self.prgFilePath.text(), 'r') as prg:
                        for line in prg:
                            if toolPattern.match(line):
                                break
                            else:
                                f.write(line)
                    curTool = None
                    threshold = self.optJudgeHolesCount.value()
                    for i, v in enumerate(self.result):
                        if (self.chkSP1.isChecked() and v[0] > threshold) or \
                                (self.chkSP2.isChecked() and v[1] > threshold) or \
                                (self.chkSP3.isChecked() and v[2] > threshold) or \
                                (self.chkSP4.isChecked() and v[3] > threshold) or \
                                (self.chkSP5.isChecked() and v[4] > threshold) or \
                                (self.chkSP6.isChecked() and v[5] > threshold):
                            hole = self.data[i]
                            pos = packPos(hole[3], hole[4])
                            if pos in self.toolPos:
                                tool = self.toolPos[pos]
                                if tool != curTool:
                                    if curTool is not None:
                                        f.write('M19\n')
                                    curTool = tool
                                    f.write(tool+'\n')
                                    f.write(self.toolDepth[tool]+'\n')
                                f.write(prgPos(hole[3], hole[4]))
                                f.write('\n')
                    if curTool is not None:
                        f.write('M19\n')
                    f.write('T00\n')
                    f.write('M30\n')
                QMessageBox.information(
                    self, '完成', '深度异常的返工背钻程序已保存完毕.', QMessageBox.Ok)
            except Exception as e:
                QMessageBox.critical(
                    self, '错误', '输出返工钻孔程序时出现错误：\n'+e.args[0], QMessageBox.Ok)

    def show_plot(self):
        if (not self.data) or (not self.toolPos):
            QMessageBox.information(
                self, '提示', '请先载入钻孔程序及背钻加工记录数据.', QMessageBox.Yes)
            return
        prg_file = self.prgFilePath.text()
        if not self.posShifted:
            shiftPos = calc_shift(prg_file, self.data)
            if shiftPos is None:
                QMessageBox.warning(
                    self, '错误', '自动获取加工记录中的坐标平移量失败，无法将加工坐标与程序坐标进行匹配！', QMessageBox.Ok)
                shiftPos = (0, 0)

            # 平移坐标数据
            for r in self.data:
                r[3] = r[3]-shiftPos[0]
                r[4] = -r[4]+shiftPos[1]

            self.lblShiftX.setText(str(shiftPos[0]))
            self.lblShiftY.setText(str(shiftPos[1]))
            self.posShifted = True

        self.lastClickedPoints = []
        self.plot.clear()

        self.result = calc_outliers(
            self.data, lag=self.optJudgeHolesCount.value(),
            tol=self.optJudgeThreshold.value()/1000
        )

        s1 = pg.ScatterPlotItem(size=self.markerSize, pen=pg.mkPen(
            None), brush=self.okBrush)
        s2 = pg.ScatterPlotItem(size=self.markerSize, pen=pg.mkPen(
            None), brush=self.ngBrush)

        ng_points = []
        ok_points = []
        threshold = self.optJudgeHolesCount.value()
        for i, v in enumerate(self.result):
            if (self.chkSP1.isChecked() and v[0] > threshold) or \
                    (self.chkSP2.isChecked() and v[1] > threshold) or \
                    (self.chkSP3.isChecked() and v[2] > threshold) or \
                    (self.chkSP4.isChecked() and v[3] > threshold) or \
                    (self.chkSP5.isChecked() and v[4] > threshold) or \
                    (self.chkSP6.isChecked() and v[5] > threshold):
                ng_points.append({
                    'pos': [self.data[i][3], self.data[i][4]],
                    'data': i
                })
            else:
                ok_points.append({
                    'pos': [self.data[i][3], self.data[i][4]],
                    'data': i
                })
        self.lblNGCount.setText(str(len(ng_points)))
        s1.addPoints(ok_points)
        s2.addPoints(ng_points)
        self.plot.addItem(s1)
        self.plot.addItem(s2)
        s1.sigClicked.connect(self.clicked)
        s2.sigClicked.connect(self.clicked)

    def create_table_item(self, value):
        if value is None:
            value = ''
        else:
            value = str(value)
        return QTableWidgetItem(value)

    def clicked(self, plot, points):
        for p in self.lastClickedPoints:
            p.resetPen()
        for p in points:
            p.setPen(self.clickedPen)
            index = p.data()
            hole = self.data[index]

            self.drillTime.setText(hole[0])
            self.drillN.setText(str(hole[1]))
            x = hole[3]
            y = hole[4]
            self.drillPosX.setText(str(round(x, 3)))
            self.drillPosY.setText(str(round(y, 3)))
            self.drillTool.setText('')
            self.drilDepth.setText('')
            pos = packPos(x, y)
            if pos in self.toolPos:
                tool = self.toolPos[pos]
                if tool in self.toolSize:
                    self.drillTool.setText(self.toolSize[tool])
                else:
                    self.drillTool.setText(tool)
                if tool in self.toolDepth:
                    self.drilDepth.setText(self.toolDepth[tool][4:])

            for row in range(6):
                for col in range(1, 4):
                    item = self.create_table_item(hole[row*3+col+4])
                    if col == 2 and self.result[index][row] > self.optJudgeHolesCount.value():
                        item.setBackground(QColor(200, 50, 50))
                    self.theHeightTable.setItem(row, col, item)
            break
        self.lastClickedPoints = points

if __name__=='__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
