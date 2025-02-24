import comtypes
import comtypes.client as client
import gc
import time
from datetime import datetime
import matplotlib.pyplot as plt
import numpy as np
import os
import sys
import csv

%matplotlib ipympl

############################Gamry Classes#################################
class GamryCOMError(Exception):
    pass

def gamry_error_decoder(e):
    if isinstance(e, comtypes.COMError):
        hresult = 2**32+e.args[0]
        if hresult & 0x20000000:
            return GamryCOMError('0x{0:08x}: {1}'.format(2**32+e.args[0], e.args[1]))
    return e

class GamryDtaqEvents(object):
    def __init__(self, dtaq):
        self.dtaq = dtaq
        self.acquired_points = [] #This is a list of tuples with 10 columns of measurement data
        
    def cook(self):
        count = 1
        while count > 0:
            count, points = self.dtaq.Cook(2**15)
            # print(count)
            self.acquired_points.extend(zip(*points))

    def _IGamryDtaqEvents_OnDataAvailable(self, this):
        self.cook()

    def _IGamryDtaqEvents_OnDataDone(self, this):
        self.cook() # final data acquisition 
        time.sleep(1)
        global active
        active = False

        print("DONE ")




GamryCOM=client.GetModule(['{BD962F0D-A990-4823-9CF5-284D1CDD9C6D}', 1, 0])
devices=client.CreateObject('GamryCOM.GamryDeviceList')
#Gamry Instrument object
pstat=client.CreateObject('GamryCOM.GamryPC6Pstat')
try:
    pstat.Init(devices.EnumSections()[0])
except IndexError:
    raise Exception("\n**ERROR - Unable to initialize. Restart the instrument then try again**\n")
#Open
pstat.Open()
pstat.SetCtrlMode(GamryCOM.PstatMode)#Set it to Potentiostat mode
#Data acquisition object cpiv
dtaqcpiv=client.CreateObject('GamryCOM.GamryDtaqCpiv')
dtaqcpiv.Init(pstat)


dtaqsink = GamryDtaqEvents(dtaqcpiv)
connection = client.GetEvents(dtaqcpiv, dtaqsink)
print("\n========================================================================")
print(devices.EnumSections()[0], " Initialization Completed In Potentiostat Mode")

# file_path = 'signal\\'
# name = 'FordProfileNoCalib2.csv'
# idx = name.find('SR')
# idx_end = name.find('k.')
# fre_sr = float(name[idx+2:idx_end])


fre_sr =100
file_path ="profile\\"

name = 'Sin_m1_F30.0k_SR100.0k.csv'
print('Samplerate: {}kHz'.format(fre_sr))

fpath = file_path +name

date = '20250120'
amp =6

target = "Cell 10 feq {}".format(fre_sr)
# "stator"
# "Resistor"

f = open(fpath)

PointsList_o = f.readlines()
numOfPoints = len(PointsList_o)
PointsList = [float(i)*amp for i in PointsList_o]
PointsList_o = [float(ii) for ii in PointsList_o]

Cycles = 1
temprate = float(1/(((fre_sr)*1000)))
temprate = round(temprate,8)
SampleRate = temprate
timeList = []
timeVal = 0

Sig=client.CreateObject('GamryCOM.GamrySignalArray')#Create Signal Object
Sig.Init(pstat, Cycles, SampleRate, numOfPoints, PointsList, GamryCOM.PstatMode)

pstat.SetSignal(Sig)
pstat.SetCell(GamryCOM.CellOn)

# dtaqsink = GamryDtaqEvents(dtaqcpiv)
# connection = client.GetEvents(dtaqcpiv, dtaqsink)


dtaqcpiv.Run(True)

# client.PumpEvents(1)

# del connection
active = True
while active == True:
    client.PumpEvents(1)
    time.sleep(0.1)
    # counter+=1
    # prograssList.append(counter)
    # if self.progressBar.value() >= 30:
    #     self.progressBar.setValue(29)
    # else:
    #     self.progressBar.setValue(counter)


pstat.SetCell(GamryCOM.CellOff)
pstat.Close()

gc.collect()

print(len(dtaqsink.acquired_points))

data_l = dtaqsink.acquired_points
data = np.array(data_l)



plt.figure(0)
plt.plot(data[:,0],PointsList)
for i in range(10):
    plt.figure(i+1)
    plt.plot(data[:,0], data[:,i])



fig, axs = plt.subplots(4,sharex=True,figsize=(18, 16))
# fig(figsize=(8, 6))
save_name = 'Pstat_{}_{}_amp{}'.format(target,name[:-4],amp)

fig.suptitle(save_name)
axs[0].plot(data[:,0], np.array(PointsList_o))
axs[0].set_title('Orignal_signal')
axs[2].plot(data[:,0], data[:,1])
axs[2].set_title('Vol')
axs[3].plot(data[:,0], data[:,3])
axs[3].set_title('Current')
axs[1].plot(data[:,0], data[:,4])
axs[1].set_title('Sendout')


cdir = os.getcwd() #Get the current directory
parent_dir = os.path.join(cdir,date)

output_data = os.path.join(parent_dir, 'Pstat')

if not os.path.exists(output_data):
    os.makedirs(output_data) 


output_fig = os.path.join(parent_dir, 'fig\Pstat')

if not os.path.exists(output_fig):
    os.makedirs(output_fig)

  

plt.savefig('{}/{}.png'.format(output_fig, save_name),bbox_inches ="tight")
np.savetxt('{}/{}.csv'.format(output_data,save_name), data, delimiter=",")

plt.show()
