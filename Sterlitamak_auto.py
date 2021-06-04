import xlrd
import xlwt
from datetime import date
from datetime import time
import os
import obspy
from obspy.core import UTCDateTime
from obspy import read
from obspy.core.stream import Stream


def del_chan(st,cha='LOG'):
    st2=Stream()
    for i in st:
        if i.stats.channel!=cha:
            st2.append(i)
    return st2

class Well():

    def __init__(self):
        self.num=0
        self.X=0
        self.Y=0
        self.Z=0
        self.H=0
        self.Az=0
        self.Fi=0
        self.Zab=0
        self.Length=0
        self.dt=0
        self.emuls=0
        self.gran=0
        self.mass=self.emuls+self.gran
        self.strmass=str(self.emuls)+'/'+str(self.gran)
        self.si='Искра-С'
        self.typeVV='Сибирит 1200/Гранулотол'


    def createwell(self,lst):
        self.num=int(lst[0].value)
        self.X=lst[1].value
        self.Y=lst[2].value
        self.Z=lst[3].value
        self.H=lst[4].value
        self.D=lst[5].value
        self.Az=lst[6].value
        self.Fi=lst[7].value
        self.Zab=lst[8].value
        self.Length=lst[9].value
        self.dt=lst[10].value
        self.emuls=int(lst[11].value)
        if lst[12].value!='':
            self.gran=int(lst[12].value)
        else:
            self.gran = 0
        self.si=lst[13].value
        self.typeVV=lst[14].value
        self.mass=self.emuls+self.gran
        self.strmass = str(int(self.emuls))+'/'+str(int(self.gran))

class Block():

    def __init__(self):
        self.num=0
        self.date=date(2000,1,1)
        self.time=time(0,0,0)
        self.wellscount=0
        self.si='Искра-С'
        self.typeVV='Сибирит 1200/Гранулотол'
        self.wells=[]

    def get_datetime(self):
        s=self.date.split('.')
        #year='20'+s[2]
        year = s[2]
        mon=s[1]
        day=s[0]
        self.date=year+'-'+mon+'-'+day
        d=year+'-'+mon+'-'+day+'T'
        s=self.time.split('-')
        h=int(s[0])-5
        m=s[1]
        s='00'
        d=d+str(h)+':'+m+':'+s
        self.datetime=d


    def readfile(self,pth):
        book = xlrd.open_workbook(pth)
        sheet = book.sheet_by_index(0)
        self.time = sheet.cell(3, 2).value
        self.date = sheet.cell(3, 1).value
        self.num = int(sheet.cell(3, 0).value)
        self.wellscount = len(sheet.col(3, start_rowx=3))
        self.get_datetime()
        well_list=[]
        VV=''
        si=''
        for i in range(self.wellscount):
            wellone = sheet.row_slice(3+i, start_colx=3, end_colx=18)
            well=Well()
            well.createwell(wellone)
            if len(VV)<len(well.typeVV):
                VV=well.typeVV
                si=well.si
            well_list.append(well)
        self.wells=well_list
        self.typeVV=VV
        self.si=si

        book.release_resources()
        del book

    def otchet(self):
        otch=[]
        otch.append(str(self.num))
        otch.append(self.date)
        otch.append(self.time)
        otch.append(str(self.wellscount))
        otch.append(self.typeVV)
        otch.append(self.si)
        return '\t'.join(otch)+'\n'

pth='y:\\Work\\Мурыськин_Алексей\\Стерлитамак\\Сырые блоки\\'
pth2='y:\\Work\\Мурыськин_Алексей\\Стерлитамак\\\Готовые блоки\\'
pth3='y:\\Work\\Мурыськин_Алексей\\Стерлитамак\\msd\\'
pth_data='r:\\data\\!region_data_streams\\shahtau\\seismograms\\2021\\04\\'

flist=os.listdir(pth)

for j in flist:
    block=Block()
    block.readfile(pth+j)

    res=xlwt.Workbook()
    sheet=res.add_sheet(str(block.num))

    u01 = pth_data + block.date + '_Shahtau_U01.msd'
    u02 = pth_data + block.date + '_Shahtau_U02.msd'
    u04 = pth_data + block.date + '_Shahtau_U04.msd'
    u07 = pth_data + block.date + '_Shahtau_U07.msd'
    u08 = pth_data + block.date + '_Shahtau_U08.msd'
    u09 = pth_data + block.date + '_Shahtau_U09.msd'

    t=UTCDateTime(block.datetime)
    tstart=t-60*10
    tend=t+60*5

    st_u01 = read(u01, starttime=tstart, endtime=tend)
    st_u02 = read(u02, starttime=tstart, endtime=tend)
    st_u04 = read(u04, starttime=tstart, endtime=tend)
    st_u07 = read(u07, starttime=tstart, endtime=tend)
    st_u08 = read(u08, starttime=tstart, endtime=tend)
    try:
        st_u09 = read(u09, starttime=tstart, endtime=tend)
    except:
        print(str(tstart))

    st_u01=del_chan(st_u01,cha='LOG')
    st_u02 = del_chan(st_u02, cha='LOG')
    st_u04 = del_chan(st_u04, cha='LOG')
    st_u07 = del_chan(st_u07, cha='LOG')
    st_u08 = del_chan(st_u08, cha='LOG')
    st_u09 = del_chan(st_u09, cha='LOG')

    u01 = pth3 + block.date + '_Shahtau_U01.msd'
    u02 = pth3 + block.date + '_Shahtau_U02.msd'
    u04 = pth3 + block.date + '_Shahtau_U04.msd'
    u07 = pth3 + block.date + '_Shahtau_U07.msd'
    u08 = pth3 + block.date + '_Shahtau_U08.msd'
    u09 = pth3 + block.date + '_Shahtau_U09.msd'

    try:
        st_u01.write(u01, format='MSEED')
    except:
        print('WARN:'+u01+'\n')

    try:
        st_u02.write(u02, format='MSEED')
    except:
        print('WARN:' + u02 + '\n')

    try:
        st_u04.write(u04, format='MSEED')
    except:
        print('WARN:'+u04+'\n')

    try:
        st_u07.write(u07, format='MSEED')
    except:
        print('WARN:'+u07+'\n')

    try:
        st_u08.write(u08, format='MSEED')
    except:
        print('WARN:'+u08+'\n')

    try:
        st_u09.write(u09, format='MSEED')
    except:
        print('WARN:' + u09 + '\n')

    try:
        with open(pth2+'res.txt','a') as f:
            f.write(block.otchet())
    except:
        with open(pth2+'res.txt','w') as f:
            f.write(block.otchet())

    for i in range(len(block.wells)):
        sheet.write(i,0,block.wells[i].num)
        sheet.write(i,1,block.wells[i].X)
        sheet.write(i, 2, block.wells[i].Y)
        sheet.write(i, 3, block.wells[i].Z)
        sheet.write(i, 4, block.wells[i].H)
        sheet.write(i, 5, block.wells[i].D)
        sheet.write(i, 6, block.wells[i].Az)
        sheet.write(i, 7, block.wells[i].Fi)
        sheet.write(i, 8, block.wells[i].Zab)
        sheet.write(i, 9, block.wells[i].Length)
        sheet.write(i, 10, block.wells[i].dt)
        sheet.write(i, 11, block.wells[i].mass)
        sheet.write(i, 12, block.wells[i].strmass)

    res.save(pth2+str(block.num)+'.xls')

