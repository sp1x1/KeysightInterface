from aifc import Error
from random import choices
import pyvisa
from pyvisa.constants import StopBits, Parity
import sys
import pandas
import openpyxl
import threading
import wx
import time
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation
import ctypes
import os
import datetime

global Halt 
Halt = False
#Extracted from the previous dataray interface program
def connectToKeySight(VoltLim = 4, ChannelAmt = 2, CurrentLim = 10, VoltRange = "20V", CurrentRange = "120mA"):
    resourceManager = pyvisa.ResourceManager()
    if 'KeySight' not in globals():
        # try:
            # This code esablishes a connection
            print(resourceManager.list_resources())
            global KeySight
            KeySight = resourceManager.open_resource('USB0::0x0957::0x4118::MY58330002::INSTR')
            print("opened resource")
            # This code sets the prerequisites for the keysight's function during testing.
            KeySight.write('CURR:RANGE R120mA, (@1)')
            KeySight.write('VOLT:RANGE R20V, (@1)')

            #This sets the Volt limit as the volt limit provided by the user on channel one.
            KeySight.write(f'VOLT:LIM {VoltLim}, (@1)')
            KeySight.write('SENS:CURR:NPLC 1, (@1)')
            KeySight.write('SENS:VOLT:NPLC 1, (@1)')

            if ChannelAmt == 2:
                # This provides the same parameters for channel Three.
                KeySight.write(f'CURR:RANGE R{CurrentRange}, (@3)')
                KeySight.write(f'VOLT:RANGE R{VoltRange}, (@3)')

                KeySight.write(f'CURR:LIM {CurrentLim}mA, (@3)')

                KeySight.write('SENS:CURR:NPLC 0, (@3)')
                KeySight.write('SENS:VOLT:NPLC 0, (@3)')

                KeySight.write('VOLT 0.1, (@3)')

            elif ChannelAmt == 3:
                # This provides the same parameters for channel two.
                KeySight.write('CURR:RANGE R120mA, (@2)')
                KeySight.write('VOLT:RANGE R20V, (@2)')

                KeySight.write(f'VOLT:LIM {VoltLim}, (@2)')
                KeySight.write('SENS:CURR:NPLC 0, (@2)')
                KeySight.write('SENS:VOLT:NPLC 0, (@2)')

                # This provides the same parameters for channel Three.
                KeySight.write(f'CURR:RANGE R{CurrentRange}, (@3)')
                KeySight.write(f'VOLT:RANGE R{VoltRange}, (@3)')

                KeySight.write(f'CURR:LIM {CurrentLim}mA, (@3)')

                KeySight.write('SENS:CURR:NPLC 0, (@3)')
                KeySight.write('SENS:VOLT:NPLC 0, (@3)')

                KeySight.write('VOLT 0.1, (@3)')
        # except:
        #     pass
def connectToArroyo():
    resourceManager = pyvisa.ResourceManager()
    if 'Arroyo' not in globals():
        global Arroyo
        Arroyo = resourceManager.open_resource('ASRL3::INSTR', baud_rate=38400, data_bits=8, flow_control=0, 
        parity=Parity.none, stop_bits=StopBits.one)
def disconnectKeySight():
    if 'KeySight' in globals():
        global KeySight
        KeySight.write('*OPC?')
        KeySight.write('OUTP 0, (@1)')
        KeySight.write('OUTP 0, (@2)')
        KeySight.write('OUTP 0, (@3)')

        del KeySight
def disconnectArroyo():
    if 'Arroyo' in globals():
        global Arroyo
        Arroyo.write('TEC:OUTPUT 0')
        del Arroyo

def checkIfInBox(p0, p1, p2, p3, p4):

    x1, y1 = p1
    x2, y2 = p2
    x3, y3 = p3
    x4, y4 = p4

    x0, y0 = p0

    if crossproduct(p0,p1,p2) and crossproduct(p0, p2, p3) and crossproduct(p0, p3, p4) and crossproduct(p0, p4, p1):
        return True
    if p0 == p3:
        return True
        
def crossproduct(p0,p1,p2, opp=False):
    x1, y1 = p1
    x2, y2 = p2
    x0, y0 = p0

    v1 = (x2-x1, y2-y1)   # Vector 1
    v2 = (x2-x0, y2-y0)   # Vector 2
    xp = v1[0]*v2[1] - v1[1]*v2[0]  # Cross product

    if xp >= 0 and opp:
        return False
    elif xp <= 0 and opp:
        return True
    elif xp >= 0:
        return True
    elif xp < 0:
        return False

class HALT(Exception): pass

def wrapFuncForGraphUpdate(CurrCH1, PhotoCurr):
    theApplication.updateGraph(CurrCH1, PhotoCurr)

def wrapFuncForsetupGraph():
    theApplication.setupGraph()

def wrapFuncForClosingGraph():
    theApplication.destroyGraph()

def checkHalt(inACurrentSweep, AllDataCollected = [], columns=["Nothing"]):
    global Halt
    if Halt:
        Halt = False
        if not inACurrentSweep:
            excelDataFrame = pandas.DataFrame(AllDataCollected, index=None, columns=columns)
            now = datetime.datetime.now()
            currentDate = now.date()
            excelDataFrame.to_excel(f"{theApplication.getFolderName()}{currentDate}'s_{theApplication.getDeviceInformation()}.xlsx")
            ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)
            disconnectKeySight()
            disconnectArroyo()
            raise HALT
        
        if inACurrentSweep:
            pass
# Before we start the loop function we define a function that allows a piece of custom text to be file ready
def makeCustomTextFileReady(text):
    newText = text
    invaildFileCharacters = ["/","\\",":","*","?",'"',"<",">","|",".","*"]
    for item in invaildFileCharacters:
        newText = newText.replace(item,"")
    return newText
def makeFolderNameReady(text):
    newText = text
    newText = newText.replace("\\", "/", 100)
    return str(newText)

def HaltFunc(e):
    global Halt
    Halt = True


def SweepCurr(strCur, stepCur, endCur, strCur2=None, stepCur2=None, endCur2=None):

    #Start currents must be equal
    args = (strCur, stepCur, endCur, strCur2, stepCur2, endCur2)
    
    go = True
    overallIterations = range(0, int((endCur-strCur)/stepCur)+1)
    for item in args:
        
        if item == None:
            _strCur2 = strCur
            _stepCur2 = stepCur
            _endCur2 = endCur
            overallIterations2 = range(0, int((_endCur2-_strCur2)/_stepCur2)+1)
        else:
            _strCur2 = strCur2
            _stepCur2 = stepCur2
            _endCur2 = endCur2
            overallIterations2 = range(0, int((_endCur2-_strCur2)/_stepCur2)+1)
        
        if endCur == strCur and stepCur == 0:
            overallIterations = [0]
        if _endCur2 == _strCur2 and _stepCur2 == 0:
            overallIterations2 = [0]
        
    
        _point1 = (strCur, _strCur2)
        _point2 = (strCur, endCur)
        _point3 = (endCur, _endCur2)
        _point4 = (endCur, _strCur2)
    try:
        if go:
            CurrentCH1 = []
            VoltCH1 = []
            CurrentCH2 = []
            VoltCH2 = []
            CurrentCH3 = []
            KeySight.write("OUTP 1, (@1)")
            KeySight.write("OUTP 1, (@2)")
            KeySight.write('*OPC?')
            time.sleep(5)
            checkHalt(True)
            
            for iter in range(0, overallIterations):
                checkHalt(True)
                SetCurrentTo = (strCur + (iter * stepCur))
                KeySight.write(f'CURR {SetCurrentTo/1000}, (@1)')
    
                for iteration in range(0, overallIterations2):
                    checkHalt(True)
                    SetCurrentTo2 = (_strCur2 + (iteration * _stepCur2))
                    KeySight.write(f'CURR {SetCurrentTo2/1000}, (@2)')
                    time.sleep(.1)
                    if checkIfInBox((SetCurrentTo, SetCurrentTo2), _point1, _point2, _point3, _point4):
                        CurrentCH1.append(SetCurrentTo)
                        VoltCH1.append(abs(float(KeySight.query('MEAS:VOLT? (@1)'))-((SetCurrentTo/1000)*8)))  
                        time.sleep(.1)
                        CurrentCH2.append(SetCurrentTo2)
                        time.sleep(.1)
                        VoltCH2.append(abs(float(KeySight.query('MEAS:VOLT? (@2)'))-((SetCurrentTo2/1000)*8)))
                        time.sleep(.1)
                        CurrentCH3.append(float(KeySight.query('MEAS:CURR? (@3)'))*1000) 
            else:
                KeySight.write("OUTP 0, (@1)")
                KeySight.write("OUTP 0, (@2)")
                AllData = zip(CurrentCH1, VoltCH1, CurrentCH2, VoltCH2, CurrentCH3)
                return list(AllData)
    except:
        pass


def setTemp(Temp):
    Arroyo.write('TEC:OUTPUT 1')
    print(f'TEC:T {Temp}')
    Arroyo.write(f'TEC:T {Temp}')
    Tolerance = .5

    TempStable = False
    while TempStable != True:
        realTemp = 0
        for iter in range(5):
            currTemp = float(Arroyo.query('TEC:T?'))
            realTemp += abs(currTemp-Temp)
            time.sleep(1)
        IsThisTolerable = realTemp/5
        realTemp = 0
        if IsThisTolerable <= Tolerance:
            TempStable = True


def OneDimensionalSweep(strCur,stepCur,endCur):
    args = (strCur, stepCur, endCur)
    
    try:
        for item in args:
            _ = float(item)
    except:
        theApplication.setStatus("Please enter all numerical values correctly.")
        args = None

    if endCur-strCur <= 0:
        theApplication.setStatus("Please set the end Current less that the Start Current")
        args = None

    try:
        if args:
            if endCur == strCur and stepCur == 0:
                overallIterations = [0]
            overallIterations = range(0, int((endCur-strCur)/stepCur)+1)
            CurrentCH1 = []
            VoltCH1 = []
            CurrentCH3 = []
            print(f"iterations: {int((endCur-strCur)/stepCur)+1}")
            KeySight.write("OUTP 1, (@1)")
            KeySight.write('*OPC?')
            time.sleep(5)
            checkHalt(True)
            for iter in range(0, overallIterations):
                checkHalt(True)
                SetCurrentTo = (strCur + (iter * stepCur))    
                KeySight.write(f'CURR {SetCurrentTo/1000}, (@1)')

                CurrentCH1.append(SetCurrentTo)
                VoltCH1.append(abs(float(KeySight.query('MEAS:VOLT? (@1)'))-((SetCurrentTo/1000)*8)))
                time.sleep(.5) 
                CurrentCH3.append(float(KeySight.query('MEAS:CURR? (@3)'))*1000)
                time.sleep(.1)
            else:
                KeySight.write("OUTP 0, (@1)")
                alldata = zip(CurrentCH1, VoltCH1, CurrentCH3)
                return list(alldata)
    except:
        pass


def BoxLoopFunction(strCur, stepCur, endCur, strTemp, stepTemp, endTemp, VoltLim, CurrLim, ChannelAmt,VoltRange,CurrRange):
    args = (strCur, stepCur, endCur, strTemp, stepTemp, endTemp, VoltLim, CurrLim, ChannelAmt)

    #This unusal function below is an anti sleep function to ensure the function runs completely.
    ctypes.windll.kernel32.SetThreadExecutionState(0x80000002)

    try:
        for item in args:
            _ = float(item)
    except:
        args = None
    try:
        if args:
            
            print("good args")
            _strCur, _stepCur, _endCur, _strTemp, _stepTemp, _endTemp, _VoltLim, _CurrLim, _ChannelAmt = int(strCur), float(stepCur), int(endCur), int(strTemp), float(stepTemp), int(endTemp), int(VoltLim), int(CurrLim), int(ChannelAmt)
            connectToKeySight(_VoltLim, _ChannelAmt, _CurrLim, VoltRange, CurrRange)
            connectToArroyo()

            overallIterations = range(0, int((_endTemp-_strTemp)/_stepTemp)+1)
            if _endTemp == _strTemp and _stepTemp == 0:
                overallIterations = [0]
            if _strCur > _endCur:
                theApplication.setStatus("Please lower the end current to be less than the start.")
                raise Error
            if _endCur >= 120:
                theApplication.setStatus("Please lower the end current, it goes out of range.")
                raise Error
            if _strTemp > _endTemp:
                theApplication.setStatus("Please lower the end tempurature to be less than the start.")
                raise Error
            if _endTemp >= 150:
                theApplication.setStatus("Please lower the end tempurature, it goes out of range.")
                raise Error
            if _VoltLim >= 20:
                theApplication.setStatus("Please Lower the volt limit to be in the range.")
                raise Error
            if _CurrLim >= 120:
                theApplication.setStatus("Please Lower the current limit to be in the range.")
                raise Error


            KeySight.write('OUTP 1, (@3)')
            KeySight.write('*OPC?')

            DataFromSweep = []
            AllDataCollected = []
            AllCurrentChannelOneDataCollected = []
            AllCurrentChannelThreeDataCollected = []

            wrapFuncForsetupGraph()

            for iter in range(0, overallIterations):
                checkHalt(False, AllDataCollected, ["Current CH1 (mA)", "Voltage CH1", 
                "Current CH2", "Voltage CH2", "Current CH3 (mA)", "Temperature (C)"])
                setTemp(_strTemp+(iter*_stepTemp))
                DataFromSweep = SweepCurr(_strCur, _stepCur, _endCur)
                CurrentCH1, VoltCH1, CurrentCH2, VoltCH2, CurrentCH3 = zip(*DataFromSweep)
                
                AllCurrentChannelOneDataCollected.extend(CurrentCH1)
                AllCurrentChannelThreeDataCollected.extend(CurrentCH3)

                wrapFuncForGraphUpdate(AllCurrentChannelOneDataCollected, AllCurrentChannelThreeDataCollected)
                Temp = []
                for i in range(0, len(CurrentCH1)-1):
                    checkHalt(False, AllDataCollected, ["Current CH1 (mA)", "Voltage CH1", 
                "Current CH2", "Voltage CH2", "Current CH3 (mA)", "Temperature (C)"])
                    Temp.append(_strTemp+(iter*_stepTemp))
                AllDataCollected.extend(zip(CurrentCH1, VoltCH1, CurrentCH2, VoltCH2, CurrentCH3, Temp))
            else:
                excelDataFrame = pandas.DataFrame(AllDataCollected, index=None, columns=["Current CH1 (mA)", "Voltage CH1", 
                "Current CH2", "Voltage CH2", "Current CH3 (mA)", "Temperature (C)"])
                now = datetime.datetime.now()
                currentDate = now.date()
                excelDataFrame.to_excel(f"{theApplication.getFolderName()}{currentDate}'s_{theApplication.getDeviceInformation()}.xlsx")

                #Make the computer able to sleep again.
                ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)

                disconnectKeySight()
                disconnectArroyo()
    except:
        pass

def StripLoopFunction(strCur, stepCur, endCur, strCur2, stepCur2, endCur2, strTemp, stepTemp, endTemp, VoltLim, CurrLim, ChannelAmt, VoltRange,CurrRange):
    args = (strCur, stepCur, endCur, strCur2, stepCur2, endCur2, strTemp, stepTemp, endTemp, VoltLim, CurrLim, ChannelAmt,)

    #This unusal function below is an anti sleep function to ensure the function runs completely.
    ctypes.windll.kernel32.SetThreadExecutionState(0x80000002)

    try:
        for item in args:
            _ = float(item)
    except:
        args = None
    
    if args:
        
        print("good args")
        _strCur, _stepCur, _endCur, _strCur2, _endCur2, _stepCur2, _strTemp, _stepTemp, _endTemp, _VoltLim, _CurrLim, _ChannelAmt = int(strCur), float(stepCur), int(endCur), int(strCur2), int(endCur2), float(stepCur2), int(strTemp), float(stepTemp), int(endTemp), int(VoltLim), int(CurrLim), int(ChannelAmt)
        connectToKeySight(_VoltLim, _ChannelAmt, _CurrLim, VoltRange, CurrRange)
        connectToArroyo()
        overallIterations = range(0, int((_endTemp-_strTemp)/_stepTemp)+1)
        if _endTemp == _strTemp and _stepTemp == 0:
            overallIterations = [0]
        if _strCur > _endCur:
                theApplication.setStatus("Please lower the end current to be less than the start.")
                raise Error
        if _endCur >= 120:
            theApplication.setStatus("Please lower the end current, it goes out of range.")
            raise Error
        if _strCur2 > _endCur2:
                theApplication.setStatus("Please lower the second end current to be less than the start.")
                raise Error
        if _endCur2 >= 120:
            theApplication.setStatus("Please lower the second end current, it goes out of range.")
            raise Error
        if _strTemp > _endTemp:
            theApplication.setStatus("Please lower the end tempurature to be less than the start.")
            raise Error
        if _endTemp >= 150:
            theApplication.setStatus("Please lower the end tempurature, it goes out of range.")
            raise Error
        if _VoltLim >= 20:
            theApplication.setStatus("Please Lower the volt limit to be in the range.")
            raise Error
        if _CurrLim >= 120:
            theApplication.setStatus("Please Lower the current limit to be in the range.")
            raise Error

        KeySight.write('OUTP 1, (@3)')
        KeySight.write('*OPC?')

        DataFromSweep = []
        AllDataCollected = []
        AllCurrentChannelOneDataCollected = []
        AllCurrentChannelThreeDataCollected = []

        wrapFuncForsetupGraph()
        
        for iter in range(0, overallIterations):
            checkHalt(False, AllDataCollected, ["Current CH1 (mA)", "Voltage CH1", 
                "Current CH2", "Voltage CH2", "Current CH3 (mA)", "Temperature (C)"])

            setTemp(_strTemp+(iter*_stepTemp))
            DataFromSweep = SweepCurr(_strCur, _stepCur, _endCur, _strCur2, _stepCur2, _endCur2)
            CurrentCH1, VoltCH1, CurrentCH2, VoltCH2, CurrentCH3 = zip(*DataFromSweep)
            AllCurrentChannelOneDataCollected.extend(CurrentCH1)
            AllCurrentChannelThreeDataCollected.extend(CurrentCH3)

            wrapFuncForGraphUpdate(AllCurrentChannelOneDataCollected, AllCurrentChannelThreeDataCollected)
            Temp = []
            for i in range(0, len(CurrentCH1)-1):
                checkHalt(False, AllDataCollected, ["Current CH1 (mA)", "Voltage CH1", 
                "Current CH2", "Voltage CH2", "Current CH3 (mA)", "Temperature (C)"])
                Temp.append(_strTemp+(iter*_stepTemp))
            AllDataCollected.extend(zip(CurrentCH1, VoltCH1, CurrentCH2, VoltCH2, CurrentCH3, Temp))
        else:
            excelDataFrame = pandas.DataFrame(AllDataCollected, index=None, columns=["Current CH1 (mA)", "Voltage CH1", 
            "Current CH2", "Voltage CH2", "Current CH3 (mA)", "Temperature (C)"])
            now = datetime.datetime.now()
            currentDate = now.date()
            excelDataFrame.to_excel(f"{theApplication.getFolderName()}{currentDate}'s_{theApplication.getDeviceInformation()}.xlsx")

            #Make the computer able to sleep again.
            ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)

            disconnectKeySight()
            disconnectArroyo()

def OneDLoopFunction(strCur, stepCur, endCur, strTemp, stepTemp, endTemp, VoltLim, CurrLim, ChannelAmt, VoltRange, CurrRange):
    args = (strCur, stepCur, endCur, strTemp, stepTemp, endTemp, VoltLim, CurrLim, ChannelAmt)

    #This unusal function below is an anti sleep function to ensure the function runs completely.
    ctypes.windll.kernel32.SetThreadExecutionState(0x80000002)

    try:
        for item in args:
            _ = float(item)
            
    except:
        args = None

    try:
        if args:
            # plt.xlabel("current CH1")
            # plt.ylabel("photo current")
            # plt.show()
            
            print("good args")
            _strCur, _stepCur, _endCur, _strTemp, _stepTemp, _endTemp, _VoltLim, _CurrLim, _ChannelAmt = int(strCur), float(stepCur), int(endCur), int(strTemp), float(stepTemp), int(endTemp), int(VoltLim), int(CurrLim), int(ChannelAmt)
            connectToKeySight(_VoltLim, _ChannelAmt, _CurrLim, VoltRange, CurrRange)
            connectToArroyo()

            overallIterations = range(0, int((_endTemp-_strTemp)/_stepTemp)+1)
            if _endTemp == _strTemp and _stepTemp == 0:
                overallIterations = [0]
            if _strCur > _endCur:
                theApplication.setStatus("Please lower the end current to be less than the start.")
                raise Error
            if _endCur >= 120:
                theApplication.setStatus("Please lower the end current, it goes out of range.")
                raise Error
            if _strTemp > _endTemp:
                theApplication.setStatus("Please lower the end tempurature to be less than the start.")
                raise Error
            if _endTemp >= 150:
                theApplication.setStatus("Please lower the end tempurature, it goes out of range.")
                raise Error
            if _VoltLim >= 20:
                theApplication.setStatus("Please Lower the volt limit to be in the range.")
                raise Error
            if _CurrLim >= 120:
                theApplication.setStatus("Please Lower the current limit to be in the range.")
                raise Error
            
            wrapFuncForsetupGraph()

            KeySight.write('OUTP 1, (@3)')
            KeySight.write('*OPC?')

            time.sleep(5)

            DataFromSweep = []
            AllDataCollected = []

            AllCurrentChannelOneDataCollected = []
            AllCurrentChannelThreeDataCollected = []
            
            for iter in range(0, overallIterations):
                checkHalt(False, AllDataCollected, ["Current CH1 (mA)", "Voltage CH1", 
                "Current CH3 (mA)", "Temperature (C)"])
                setTemp(_strTemp+(iter*_stepTemp))
                DataFromSweep = OneDimensionalSweep(_strCur, _stepCur, _endCur)
                Curr1, Voltage1, Curr3 = zip(*DataFromSweep)
                AllCurrentChannelOneDataCollected.extend(Curr1)
                AllCurrentChannelThreeDataCollected.extend(Curr3)

                wrapFuncForGraphUpdate(AllCurrentChannelOneDataCollected, AllCurrentChannelThreeDataCollected)
                # plt.cla()
                # plt.plot(Curr1,Curr3)
                Temp = []
                for i in range(0, len(Curr1)-1):
                    checkHalt(False, AllDataCollected, ["Current CH1 (mA)", "Voltage CH1", 
                "Current CH3 (mA)", "Temperature (C)"])
                    Temp.append(_strTemp+(iter*_stepTemp))
                AllDataCollected.extend(zip(Curr1, Voltage1, Curr3, Temp))
                
            else:
                excelDataFrame = pandas.DataFrame(AllDataCollected, index=None, columns=["Current CH1 (mA)", "Voltage CH1", 
                "Current CH3 (mA)", "Temperature (C)"])
                now = datetime.datetime.now()
                currentDate = now.date()
                excelDataFrame.to_excel(f"{theApplication.getFolderName()}{currentDate}'s_{theApplication.getDeviceInformation()}.xlsx")
                KeySight.write('OUTP 0, (@3)')

                #Make the computer able to sleep again.
                ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)

                disconnectKeySight()
                disconnectArroyo()
    except:
        pass



class theApplication(wx.App):
    def __init__(self, redirect=False, filename=None):
        wx.App.__init__( self, redirect, filename )
        self.frameMain = wx.Frame( parent=None, id=wx.ID_ANY, size=(800,400), 
                              title='Python Interface to KeySight')
        self.pMain = wx.Panel(self.frameMain, id=wx.ID_ANY)
        self.frameMain.Show()

        '''MAIN FRAME'''
        # Test Start Button
        self.TestBeginButton = wx.Button(self.pMain, label="Begin Test", pos=(160, 50))
        self.TestBeginButton.Bind(wx.EVT_BUTTON, self.wrapperFunction)

        # Box Mode Button
        self.BoxModeButton = wx.Button(self.pMain, label="Box Mode", pos=(10,10))
        self.BoxModeButton.Bind(wx.EVT_BUTTON, self.showBoxMode)

        # Strip Mode Button
        self.StripModeButton = wx.Button(self.pMain, label="Strip Mode", pos=(160,10))
        self.StripModeButton.Bind(wx.EVT_BUTTON, self.showStripMode)

        # 1D Mode Button
        self.OneDModeButton = wx.Button(self.pMain, label="1D Mode", pos=(310,10))
        self.OneDModeButton.Bind(wx.EVT_BUTTON, self.showOneDMode)

        # Connection Controls
        self.TurnOnWithVoltageLimitHeader = wx.StaticText(self.pMain, label="Setup the Keysight's Volt Limit on Both Channels to: ____ (V) ", pos=(300, 100))
        self.TurnOnWithVoltageLimit = wx.TextCtrl(self.pMain, pos=(350,125), size=(100,-1))

        self.TurnOnWithCurrentLimitHeader = wx.StaticText(self.pMain, label="Setup the Keysight's Current Limit on the Third Channel to: ___ (mA) ", pos=(300, 150))
        self.TurnOnWithCurrentLimit = wx.TextCtrl(self.pMain, pos=(350, 175), size=(100,-1))

        self.RangeHeader = wx.StaticText(self.pMain, label="Setup the range for the third channel:", pos=(300, 225))
        self.CurrentRange = wx.ComboBox(self.pMain, value="Current Range", choices=["1uA","10uA","100uA","1mA","10mA","120mA"], pos=(300, 250), size=(110,-1))
        self.VoltRange = wx.ComboBox(self.pMain, value="Volt Range", choices=["2V","20V"], pos=(300, 275), size=(100,-1))

        #Halt Button
        self.HaltButton = wx.Button(self.pMain, label="HALT", pos=(160, 75))
        self.HaltButton.Bind(wx.EVT_BUTTON, HaltFunc)

        #File saving widgets.
        self.FileSideDisplay = wx.StaticText(self.pMain, label="Folder Location:", pos=(30, 200))
        self.FolderDestination = wx.TextCtrl(self.pMain, value="", pos=(120, 200), size=(75, -1))
        self.BrowseButton =  wx.Button(self.pMain, label="Browse", pos=(200, 200))
        self.BrowseButton.Bind(wx.EVT_BUTTON, self.Browse)

        self.CustomFileTextHeader = wx.StaticText(self.pMain, label="File Name:", pos=(120,250))
        self.CustomFileText = wx.TextCtrl(self.pMain, value="", pos=(110, 275), size=(150, -1))

        #Status Display
        self.Status = wx.StaticText(self.pMain, label="Status: Nothing to report.", pos=(120, 150))

    def showBoxMode(self, e):
        ChannelAmount = 3
        self.selection = "Box"
        self.BoxWindow = BoxFrame(title="Python Interface to Keysight (Box Mode)")
        self.BoxWindow.Show()

    def showStripMode(self, e):
        ChannelAmount = 3
        self.selection = "Strip"
        self.StripWindow = StripFrame(title="Python Interface to Keysight (Strip Mode)")
        self.StripWindow.Show()
    def showOneDMode(self, e):
        self.selection = "1D"
        ChannelAmount = 2
        self.OneDWindow = OneDFrame(title="Python Interface to Keysight (1D Mode)")
        self.OneDWindow.Show()
    
    def setStatus(self, status="Nothing to report."):
        self.Status.Label = f"Status: {status}"
    
    def Browse(self, e):
        # Creates a file dialog
        dlg = wx.DirDialog (None, "Choose input directory", "",
                    wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST)
        dlg.ShowModal()
        self.FolderDestination.Value= str(dlg.GetPath())
        dlg.Destroy()
    def getFolderName(self):
        # Make all the back slashes foward slashes.
        FolderText = makeFolderNameReady(self.FolderDestination.Value)
        # Does this folder path exist?
        if os.path.isdir(FolderText) == True:
            #Check if the last character is a slash
            if FolderText[-1] == "/":
                return FolderText
            else:
                return FolderText+"/"
        else:
            return ""

    def getDeviceInformation(self):
        # The aim of this function is to provide the file name with the device information.
        if self.CustomFileText:
            return makeCustomTextFileReady(self.CustomFileText)
        else:
            return ""
    def setupGraph(self):
        plt.xlabel("Current CH1")
        plt.ylabel("Photo current")
        plt.show()

    def updateGraph(self, CurrCH1, PhotoCurr):

        plt.cla()
        plt.plot(CurrCH1, PhotoCurr)
    
    def destroyGraph(self):
        plt.close()
        
    def wrapperFunction(self, e):
        selection = self.selection
        if selection == "Box":
            print("we doing this")
            _arguments = (self.BoxWindow.StartingDriveCurrent.Value, self.BoxWindow.DriveCurrentStep.Value, 
            self.BoxWindow.DriveCurrentEnd.Value, self.BoxWindow.StartingTemp.Value, self.BoxWindow.TempStep.Value, 
            self.BoxWindow.EndingTemp.Value, self.TurnOnWithVoltageLimit.Value, self.TurnOnWithCurrentLimit.Value, 3, self.VoltRange.Value, self.CurrentRange.Value)
            thread = threading.Thread(target=BoxLoopFunction, args=_arguments)
            thread.start()

        if selection == "1D":
            print("we doing this")
            _arguments = (self.OneDWindow.StartingDriveCurrent.Value, self.OneDWindow.DriveCurrentStep.Value, 
            self.OneDWindow.DriveCurrentEnd.Value, self.OneDWindow.StartingTemp.Value, self.OneDWindow.TempStep.Value, 
            self.OneDWindow.EndingTemp.Value, self.TurnOnWithVoltageLimit.Value, self.TurnOnWithCurrentLimit.Value, 2, self.VoltRange.Value, self.CurrentRange.Value)
            thread = threading.Thread(target=OneDLoopFunction, args=_arguments)
            thread.start()

        if selection == "Strip":
            print("we doing this")
            _arguments = (self.StripWindow.StartingDriveCurrent.Value, self.StripWindow.DriveCurrentStep.Value, 
            self.StripWindow.DriveCurrentEnd.Value, self.StripWindow.StartingDriveCurrent2.Value, self.StripWindow.DriveCurrentStep2.Value, 
            self.StripWindow.DriveCurrentEnd2.Value, self.StripWindow.StartingTemp.Value, self.StripWindow.TempStep.Value, 
            self.StripWindow.EndingTemp.Value, self.TurnOnWithVoltageLimit.Value, self.TurnOnWithCurrentLimit.Value, 2, self.VoltRange.Value, self.CurrentRange.Value)
            thread = threading.Thread(target=StripLoopFunction, args=_arguments)
            thread.start()


class BoxFrame(wx.Frame):
    def __init__(self, title, parent=None, id=wx.ID_ANY, size=(450,300)):
        wx.Frame.__init__(self, parent=parent, title=title, id=id,size=size)

        self.pBox = wx.Panel(self, wx.ID_ANY)
         # Current setting on channel one during test.
        self.StartingDriveCurrentHeader = wx.StaticText(self.pBox, label = "Starting Current CH1 & 2(mA)", pos=(10,10))
        self.StartingDriveCurrent = wx.TextCtrl(self.pBox, pos=(10,35), size=(100,-1))
        
        self.DriveCurrentStepHeader = wx.StaticText(self.pBox, label = "Current Step CH1 & 2 ", pos=(10,60))
        self.DriveCurrentStep = wx.TextCtrl(self.pBox, pos=(10,85), size=(100,-1))

        self.DriveCurrentEndHeader = wx.StaticText(self.pBox, label = "Ending Current CH1 & 2 (mA)", pos=(10,110))
        self.DriveCurrentEnd = wx.TextCtrl(self.pBox, pos=(10,135), size=(100,-1))

        # # Current setting on channel two during test.
        # self.StartingDriveCurrentHeader2 = wx.StaticText(self.pBox, label = "Starting Current CH2 (mA)", pos=(160,10))
        # self.StartingDriveCurrent2 = wx.TextCtrl(self.pBox, pos=(160,35), size=(100,-1))
        
        # self.DriveCurrentStepHeader2 = wx.StaticText(self.pBox, label = "Current Step CH2 ", pos=(160,60))
        # self.DriveCurrentStep2 = wx.TextCtrl(self.pBox, pos=(160,85), size=(100,-1))

        # self.DriveCurrentEndHeader2 = wx.StaticText(self.pBox, label = "Ending Current CH2 (mA)", pos=(160,110))
        # self.DriveCurrentEnd2 = wx.TextCtrl(self.pBox, pos=(160,135), size=(100,-1))

        # Tempurature Setting
        self.StartingTempHeader = wx.StaticText(self.pBox, label = "Starting Temp (C)", pos=(160,10))
        self.StartingTemp = wx.TextCtrl(self.pBox, pos=(160,35), size=(100,-1))
        
        self.TempStepHeader = wx.StaticText(self.pBox, label = "Temp Step", pos=(160,60))
        self.TempStep = wx.TextCtrl(self.pBox, pos=(160,85), size=(100,-1))

        self.EndingTempHeader = wx.StaticText(self.pBox, label = "Ending Temp (C)", pos=(160,110))
        self.EndingTemp = wx.TextCtrl(self.pBox, pos=(160,135), size=(100,-1))

class StripFrame(wx.Frame):
    def __init__(self, title, parent=None):
        wx.Frame.__init__(self, parent=parent, title=title, id=wx.ID_ANY, size=(500,300))

        self.pStrip = wx.Panel(self, wx.ID_ANY)

         # Current setting on channel one during test.
        self.StartingDriveCurrentHeader = wx.StaticText(self.pStrip, label = "Starting Current CH1 (mA)", pos=(10,10))
        self.StartingDriveCurrent = wx.TextCtrl(self.pStrip, pos=(10,35), size=(100,-1))
        
        self.DriveCurrentStepHeader = wx.StaticText(self.pStrip, label = "Current Step CH1 ", pos=(10,60))
        self.DriveCurrentStep = wx.TextCtrl(self.pStrip, pos=(10,85), size=(100,-1))

        self.DriveCurrentEndHeader = wx.StaticText(self.pStrip, label = "Ending Current CH1 (mA)", pos=(10,110))
        self.DriveCurrentEnd = wx.TextCtrl(self.pStrip, pos=(10,135), size=(100,-1))

        # Current setting on channel two during test.
        self.StartingDriveCurrentHeader2 = wx.StaticText(self.pStrip, label = "Starting Current CH2 (mA)", pos=(160,10))
        self.StartingDriveCurrent2 = wx.TextCtrl(self.pStrip, pos=(160,35), size=(100,-1))
        
        self.DriveCurrentStepHeader2 = wx.StaticText(self.pStrip, label = "Current Step CH2 ", pos=(160,60))
        self.DriveCurrentStep2 = wx.TextCtrl(self.pStrip, pos=(160,85), size=(100,-1))

        self.DriveCurrentEndHeader2 = wx.StaticText(self.pStrip, label = "Ending Current CH2 (mA)", pos=(160,110))
        self.DriveCurrentEnd2 = wx.TextCtrl(self.pStrip, pos=(160,135), size=(100,-1))

        # Tempurature Setting
        self.StartingTempHeader = wx.StaticText(self.pStrip, label = "Starting Temp (C)", pos=(310,10))
        self.StartingTemp = wx.TextCtrl(self.pStrip, pos=(310,35), size=(100,-1))
        
        self.TempStepHeader = wx.StaticText(self.pStrip, label = "Temp Step", pos=(310, 60))
        self.TempStep = wx.TextCtrl(self.pStrip, pos=(310,85), size=(100,-1))

        self.EndingTempHeader = wx.StaticText(self.pStrip, label = "Ending Temp (C)", pos=(310,110))
        self.EndingTemp = wx.TextCtrl(self.pStrip, pos=(310,135), size=(100,-1))

class OneDFrame(wx.Frame):
    def __init__(self, title, parent=None):
        wx.Frame.__init__(self, parent=parent, title=title, id=wx.ID_ANY, size=(500,300))

        self.p1D = wx.Panel(self, wx.ID_ANY)

         # Current setting on channel one during test.
        self.StartingDriveCurrentHeader = wx.StaticText(self.p1D, label = "Starting Current CH1 (mA)", pos=(10,10))
        self.StartingDriveCurrent = wx.TextCtrl(self.p1D, pos=(10,35), size=(100,-1))
        
        self.DriveCurrentStepHeader = wx.StaticText(self.p1D, label = "Current Step CH1 ", pos=(10,60))
        self.DriveCurrentStep = wx.TextCtrl(self.p1D, pos=(10,85), size=(100,-1))

        self.DriveCurrentEndHeader = wx.StaticText(self.p1D, label = "Ending Current CH1 (mA)", pos=(10,110))
        self.DriveCurrentEnd = wx.TextCtrl(self.p1D, pos=(10,135), size=(100,-1))

        # # Current setting on channel two during test.
        # self.StartingDriveCurrentHeader2 = wx.StaticText(self.p1D, label = "Starting Current CH2 (mA)", pos=(160,10))
        # self.StartingDriveCurrent2 = wx.TextCtrl(self.p1D, pos=(160,35), size=(100,-1))
        
        # self.DriveCurrentStepHeader2 = wx.StaticText(self.p1D, label = "Current Step CH2 ", pos=(160,60))
        # self.DriveCurrentStep2 = wx.TextCtrl(self.p1D, pos=(160,85), size=(100,-1))

        # self.DriveCurrentEndHeader2 = wx.StaticText(self.p1D, label = "Ending Current CH2 (mA)", pos=(160,110))
        # self.DriveCurrentEnd2 = wx.TextCtrl(self.p1D, pos=(160,135), size=(100,-1))

        # Tempurature Setting
        self.StartingTempHeader = wx.StaticText(self.p1D, label = "Starting Temp (C)", pos=(160,10))
        self.StartingTemp = wx.TextCtrl(self.p1D, pos=(160,35), size=(100,-1))
        
        self.TempStepHeader = wx.StaticText(self.p1D, label = "Temp Step", pos=(160,60))
        self.TempStep = wx.TextCtrl(self.p1D, pos=(160,85), size=(100,-1))

        self.EndingTempHeader = wx.StaticText(self.p1D, label = "Ending Temp (C)", pos=(160,110))
        self.EndingTemp = wx.TextCtrl(self.p1D, pos=(160,135), size=(100,-1))

class DataFrame(wx.Frame):
    """
    Class used for creating frames other than the main one
    """
    def __init__(self, title, parent=None):
        wx.Frame.__init__(self, parent=parent, title=title)
        self.Show()


if __name__ == "__main__":
    app = theApplication()
    app.MainLoop()
    #Exit after the app is closed!
    sys.exit()

