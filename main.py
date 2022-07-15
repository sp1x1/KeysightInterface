import pyvisa
from pyvisa.constants import StopBits, Parity
import sys
import pandas
import openpyxl
import threading
import wx
import time

#Extracted from the previous dataray interface program
def connectToKeySight(VoltLim = 4, ChannelAmt = 1, CurrentLim = 10):
    resourceManager = pyvisa.ResourceManager()
    if 'KeySight' not in globals():
        try:
            # This code esablishes a connection
            global KeySight
            KeySight = resourceManager.open_resource('USB0::0x0957::0x4118::MY58330002::INSTR')

            # This code sets the prerequisites for the keysight's function during testing.
            KeySight.write('CURR:RANGE R120mA, (@1)')
            KeySight.write('VOLT:RANGE R20V, (@1)')

            #This sets the Volt limit as the volt limit provided by the user on channel one.
            KeySight.write(f'VOLT:LIM {VoltLim}, (@1)')
            KeySight.write('SENS:CURR:NPLC 1, (@1)')
            KeySight.write('SENS:VOLT:NPLC 1, (@1)')

            if ChannelAmt == 2:
                # This provides the same parameters for channel two.
                KeySight.write('CURR:RANGE R120mA, (@3)')
                KeySight.write('VOLT:RANGE R20V, (@3)')

                KeySight.write(f'CURR:LIM {CurrentLim}, (@3)')

                KeySight.write(f'VOLT: .1')
                KeySight.write('SENS:CURR:NPLC 1, (@3)')
                KeySight.write('SENS:VOLT:NPLC 1, (@3)')
            elif ChannelAmt == 3:
                # This provides the same parameters for channel two.
                KeySight.write('CURR:RANGE R120mA, (@2)')
                KeySight.write('VOLT:RANGE R20V, (@2)')

                KeySight.write(f'VOLT:LIM {VoltLim}, (@2)')
                KeySight.write('SENS:CURR:NPLC 1, (@2)')
                KeySight.write('SENS:VOLT:NPLC 1, (@2)')

            KeySight.write('SENS:CURR:NPLC 1, (@3)')
        except:
            pass
def connectToArroyo():
    resourceManager = pyvisa.ResourceManager
    if 'Arroyo' not in globals():
        global Arroyo
        Arroyo = resourceManager.open_resource('ASRL4::INSTR', baud_rate=38400, databits=8, flow_control=0, 
        parity=Parity.none, stop_bits=StopBits.one)
def disconnectKeySight():
    if 'KeySight' in globals():
        KeySight.write('OUTP 0, (@1)')
        KeySight.write('OUTP 0, (@2)')
        KeySight.write('OUTP 0, (@3)')

        del KeySight
def disconnectArroyo():
    if 'Arroyo' in globals():
        Arroyo.write('TEC:OUTPUT 0')
        del Arroyo

def checkIfInBox(p0,p1,p2,p3,p4):

    x1, y1 = p1
    x2, y2 = p2
    x3, y3 = p3
    x4, y4 = p4

    x0, y0 = p0

    if crossproduct(p0,p1,p2) and crossproduct(p0, p2, p3) and crossproduct(p0, p3, p4) and crossproduct(p0, p4, p1):
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
    elif xp < 0 and opp:
        return True
    elif xp >= 0:
        return True
    elif xp < 0:
        return False
    


def SweepCurr(strCur, stepCur, endCur, strCur2=None, stepCur2=None, endCur2=None, point1=("-1","-1"), point2=("-1","-1"), point3=("-1","-1"), point4=("-1","-1")):

    #Start currents must be equal
    args = (strCur, stepCur, endCur, strCur2, stepCur2, endCur2, point1, point2, point3, point4)
    pointargs = (point1, point2, point3, point4)
    setupPointargs = False
    go = True
    for item in args:
        try:
            _ = int(item)

            if type(item) == tuple or type(item) == list:
                try:
                    for i in item:
                        _ = int(i)
                except:
                    go = False
            if item == None:
                _strCur2 = strCur
                _stepCur2 = stepCur
                _endCur2 = endCur
        except:
            go = False
            pass
    for item in pointargs:
        for i in item:
            if i == -1:
                _point1 = (strCur, _strCur2)
                _point2 = (strCur, endCur)
                _point3 = (endCur, _endCur2)
                _point4 = (endCur, strCur)
            if i != -1:
                _point1 = point1
                _point2 = point2
                _point3 = point3
                _point4 = point4
    if go:
        CurrentCH1 = []
        VoltCH1 = []
        CurrentCH2 = []
        VoltCH2 = []
        CurrentCH3 = []
        for iter in range(0, (endCur-strCur)/stepCur):
            SetCurrentTo = (strCur + (iter * stepCur))
            SetCurrentTo2 = (_strCur2 + (iteration * _stepCur2))
            KeySight.write(f'CURR {SetCurrentTo/1000}, (@1)')

            time.sleep(.1)
            if checkIfInBox((SetCurrentTo,SetCurrentTo2),_point1,_point2,_point3,_point4):
                    CurrentCH1.append(SetCurrentTo)
                    VoltCH1.append(KeySight.query('MEAS:VOLT? (1@)'))  
            for iteration in range(0, (_endCur2-_strCur2)/_stepCur2):
                SetCurrentTo2 = (_strCur2 + (iteration * _stepCur2))
                KeySight.write(f'CURR {SetCurrentTo/1000}, (@2)')
                time.sleep(.1)
                if checkIfInBox((SetCurrentTo,SetCurrentTo2),_point1, _point2, _point3, _point4):
                    CurrentCH3.append(KeySight.query("MEAS:CURR? (@3)"))
                    time.sleep(.1)
                    CurrentCH2.append(SetCurrentTo)
                    time.sleep(.1)
                    VoltCH2.append(KeySight.query('MEAS:VOLT? (2@)'))
        else:
            AllData = zip(CurrentCH1, VoltCH1, CurrentCH2, VoltCH2, CurrentCH3)
            return list(AllData)

def setTemp(Temp):
    Arroyo.write('TEC:OUTPUT 1')
    Arroyo.write(f'TEC:T{Temp}')
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



def BoxLoopFunction(strCur, stepCur, endCur, strTemp, stepTemp, endTemp):
    args = (strCur, stepCur, endCur, strTemp, stepTemp, endTemp)
    try:
        for item in args:
            _ = int(item)
    except:
        args = None
    
    if args:
        DataFromSweep = []
        AllDataCollected = []
        for iter in range(0, (endTemp-strTemp)/stepTemp):
            setTemp(strTemp+(iter*stepTemp))
            strTemp+(iter*stepTemp)
            DataFromSweep = SweepCurr(strCur, stepCur, endCur)
            AllDataCollected.extend(zip(SweepCurr, strTemp+(iter*stepTemp)))
        else:
            excelDataFrame = pandas.DataFrame(AllDataCollected, index=None, columns=["Current CH1 (mA)", "Voltage CH1", 
            "Current CH2", "Voltage CH2", "Current CH3 (mA)", "Temperature (F)"])
            excelDataFrame.to_excel("output.xlsx")
            disconnectKeySight()
            disconnectArroyo()




class theApplication(wx.App):
    def __init__(self, redirect=False, filename=None):
        wx.App.__init__( self, redirect, filename )
        self.frameMain = wx.Frame( parent=None, id=wx.ID_ANY, size=(800,400), 
                              title='Python Interface to KeySight')
        self.pMain = wx.Panel(self.frameMain, id=wx.ID_ANY)
        self.frameMain.Show()

        '''MAIN FRAME'''
        # Test Start Button
        self.TestBeginButton = wx.Button(self.pMain, label="Begin Test", pos=(160, 35))

        # Box Mode Button
        self.BoxModeButton = wx.Button(self.pMain, label="Box Mode", pos=(10,10))
        self.BoxModeButton.Bind(wx.EVT_BUTTON, self.showBoxMode)

        # Strip Mode Button
        self.StripModeButton = wx.Button(self.pMain, label="Strip Mode", pos=(160,10))
        self.StripModeButton.Bind(wx.EVT_BUTTON, self.showStripMode)

        # 1D Mode Button
        self.OneDModeButton = wx.Button(self.pMain, label="1D Mode", pos=(310,10))
        self.OneDModeButton.Bind(wx.EVT_BUTTON, self.showOneDMode)

    def showBoxMode(self, e):
        self.BoxWindow = BoxFrame(title="Python Interface to Keysight (Box Mode)")
        self.BoxWindow.Show()

    def showStripMode(self, e):
        self.StripWindow = StripFrame(title="Python Interface to Keysight (Strip Mode)")
        self.StripWindow.Show()
    def showOneDMode(self, e):
        self.OneDWindow = OneDFrame(title="Python Interface to Keysight (1D Mode)")
        self.OneDWindow.Show()

    def wrapperFunction(self, e, selection):
        if selection == "Box":
            _arguments = (self.BoxWindow.StartingDriveCurrent.Value, self.BoxWindow.DriveCurrentStep.Value, 
            self.BoxWindow.DriveCurrentEnd.Value, self.BoxWindow.StartingTemp, self.BoxWindow.TempStep.Value, 
            self.BoxWindow.EndingTemp.Value)
            thread = threading.Thread(target=BoxLoopFunction, args=_arguments)
            thread.start()
    # If selection == ""


class BoxFrame(wx.Frame):
    def __init__(self, title, parent=None, id=wx.ID_ANY, size=(450,300)):
        wx.Frame.__init__(self, parent=parent, title=title, id=id,size=size)

        self.pBox = wx.Panel(self, wx.ID_ANY)
         # Current setting on channel one during test.
        self.StartingDriveCurrentHeader = wx.StaticText(self.pBox, label = "Starting Current CH1 (mA)", pos=(10,10))
        self.StartingDriveCurrent = wx.TextCtrl(self.pBox, pos=(10,35), size=(100,-1))
        
        self.DriveCurrentStepHeader = wx.StaticText(self.pBox, label = "Current Step CH1 ", pos=(10,60))
        self.DriveCurrentStep = wx.TextCtrl(self.pBox, pos=(10,85), size=(100,-1))

        self.DriveCurrentEndHeader = wx.StaticText(self.pBox, label = "Ending Current CH1 (mA)", pos=(10,110))
        self.DriveCurrentEnd = wx.TextCtrl(self.pBox, pos=(10,135), size=(100,-1))

        # # Current setting on channel two during test.
        # self.StartingDriveCurrentHeader2 = wx.StaticText(self.pBox, label = "Starting Current CH2 (mA)", pos=(160,10))
        # self.StartingDriveCurrent2 = wx.TextCtrl(self.pBox, pos=(160,35), size=(100,-1))
        
        # self.DriveCurrentStepHeader2 = wx.StaticText(self.pBox, label = "Current Step CH2 ", pos=(160,60))
        # self.DriveCurrentStep2 = wx.TextCtrl(self.pBox, pos=(160,85), size=(100,-1))

        # self.DriveCurrentEndHeader2 = wx.StaticText(self.pBox, label = "Ending Current CH2 (mA)", pos=(160,110))
        # self.DriveCurrentEnd2 = wx.TextCtrl(self.pBox, pos=(160,135), size=(100,-1))

        # Tempurature Setting
        self.StartingTempHeader = wx.StaticText(self.pBox, label = "Starting Temp (F)", pos=(160,10))
        self.StartingTemp = wx.TextCtrl(self.pBox, pos=(160,35), size=(100,-1))
        
        self.TempStepHeader = wx.StaticText(self.pBox, label = "Temp Step", pos=(160,60))
        self.TempStep = wx.TextCtrl(self.pBox, pos=(160,85), size=(100,-1))

        self.EndingTempHeader = wx.StaticText(self.pBox, label = "Ending Temp (F)", pos=(160,110))
        self.EndingTemp = wx.TextCtrl(self.pBox, pos=(160,135), size=(100,-1))

class StripFrame(wx.Frame):
    def __init__(self, title, parent=None):
        wx.Frame.__init__(self, parent=parent, title=title, id=wx.ID_ANY, size=(500,300))

        self.pStrip = wx.Panel(self, wx.ID_ANY)

class OneDFrame(wx.Frame):
    def __init__(self, title, parent=None):
        wx.Frame.__init__(self, parent=parent, title=title, id=wx.ID_ANY, size=(500,300))

        self.p1D = wx.Panel(self, wx.ID_ANY)

class DataFrame(wx.Frame):
    """
    Class used for creating frames other than the main one
    """
    def __init__(self, title, parent=None):
        wx.Frame.__init__(self, parent=parent, title=title)
        self.Show()


if __name__ == "__main__":
    print(checkIfInBox((2,2), (1,1), (.5,3), (4,4), (4,2)))
    app = theApplication()
    app.MainLoop()
    #Exit after the app is closed!
    sys.exit()


