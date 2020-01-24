import time
import datetime as dt
import xlrd
import scipy.stats
import math
import os
from tempfile import TemporaryFile
import xlwt

b1=xlwt.Workbook()
sheetsecond=b1.add_sheet("p227sec")
''''
splitting the data into time intervals and for every 10sec the loop is repeating and taking the values in the form of 227sec
'''


ti = []
t = dt.timedelta(0, 0, 0, 0, 0, 8)
ti.append(t)
for x in range(143):
    t = t + dt.timedelta(seconds=227)
    ti.append(t)

    ''''
    reading the directory file and giving the paths of the 54 files  and printing the users that are being comparing in the file
    '''



l=os.listdir("C:/Users/sneha/Desktop/inf_files")
for a in range(0,54):
    for b in range(0,54):
        file1="C:/Users/sneha/Desktop/inf_files/"+l[a]
        file2="C:/Users/sneha/Desktop/inf_files/"+l[b]
        print("user {} --> user {} ".format(l[a],l[b]))

        w1=xlrd.open_workbook(file1)
        w2=xlrd.open_workbook(file2)
        st=w1.sheet_by_index(0)
        st1=w2.sheet_by_index(0)
        '''
        printing the length of rows and creating the lists for doctets, realfirstpacket and duration values 
        '''
        leng = st.nrows
        docs = []
        rf1 = []
        drval = []
        ''' checking from the dates 4 to 16 and excluding the weekends with dates 9 and 10 '''
        for i in range(4, 16):
            if (i == 9 or i == 10):
                continue
            for j in range(1, leng):
                k = int(st.cell_value(j, 5))
                date = dt.datetime.fromtimestamp(k / 1000).day
                if date != i:
                    continue
                if st.cell_value(j, 9) == 0:
                    continue
                hr = dt.datetime.fromtimestamp(st.cell_value(j, 5) / 1000).hour
                '''' checking the values between 8 and 5'''
                if hr > 7 and hr <= 17:
                    docs.append(st.cell_value(j, 3))
                    rf1.append(st.cell_value(j, 5))
                    drval.append(st.cell_value(j, 9))


        doc = []
        rf = []
        dr = []
        for x in range(4, 16):
            if (x == 9 or x == 10):
                continue
            for y in range(1, st1.nrows):
                z = int(st1.cell_value(y, 5))
                dat = dt.datetime.fromtimestamp(z / 1000).day
                if dat != x:
                    continue
                if st1.cell_value(y, 9) == 0:
                    continue
                hr = dt.datetime.fromtimestamp(st1.cell_value(y, 5) / 1000).hour
                if hr > 7 and hr < 17:
                    doc.append(st1.cell_value(y, 3))
                    rf.append(st1.cell_value(y, 5))
                    dr.append(st1.cell_value(y, 9))





        def averg(l):
            return sum(l) / len(l)


        week1a = []
        week2a = []
        week1b = []
        week2b = []

        avg = []
        '''every value is being checked in the 227sec interval range and that values are being doctects / dur val calculated'''

        for k in range(4, 16):

            if k == 9 or k == 10:
                continue
            for i in range(0, 143):
                start = ti[i]
                end = ti[i + 1]
                for x in range(0, len(docs)):
                    if k != dt.datetime.fromtimestamp(rf1[x] / 1000).day:
                        continue
                    h = dt.datetime.fromtimestamp(rf1[x] / 1000).hour
                    m = dt.datetime.fromtimestamp(rf1[x] / 1000).minute
                    s = dt.datetime.fromtimestamp(rf1[x] / 1000).second
                    ms = dt.datetime.fromtimestamp(rf1[x] / 1000).microsecond
                    given = dt.timedelta(0, s, ms, 0, m, h)
                    if given >= start and given < end:
                        avg.append(docs[x] / drval[x])
                        '''if the average value is not zero the append average otherwise zero'''
                if len(avg) != 0:
                    if k >= 4 and k <= 8:
                        week1a.append(averg(avg))
                    else:
                        if k >= 11 and k <= 15:
                            week2a.append(averg(avg))
                else:
                    if k >= 4 and k <= 8:
                        week1a.append(0)
                    else:
                        if k >= 11 and k <= 15:
                            week2a.append(0)
                avg.clear()
        avg.clear()
        '''every value is being checked in the 227sec interval range and that values are being doctects / dur val calculated'''
        for k in range(4, 16):

            if k == 9 or k == 10:
                continue
            for i in range(0, 143):
                start = ti[i]
                end = ti[i + 1]
                for x in range(0, len(doc)):
                    if k != dt.datetime.fromtimestamp(rf[x] / 1000).day:
                        continue
                    h = dt.datetime.fromtimestamp(rf[x] / 1000).hour
                    m = dt.datetime.fromtimestamp(rf[x] / 1000).minute
                    s = dt.datetime.fromtimestamp(rf[x] / 1000).second
                    ms = dt.datetime.fromtimestamp(rf[x] / 1000).microsecond
                    given = dt.timedelta(0, s, ms, 0, m, h)
                    if given >= start and given < end:
                        avg.append(doc[x] / dr[x])
                        '''if the average value is not zero the append average otherwie zero'''
                if len(avg) != 0:
                    if k >= 4 and k <= 8:
                        week1b.append(averg(avg))
                    else:
                        if k >= 11 and k <= 15:
                            week2b.append(averg(avg))
                else:
                    if k >= 4 and k <= 8:
                        week1b.append(0)
                    else:
                        if k >= 11 and k <= 15:
                            week2b.append(0)
                avg.clear()

        print(len(week1a))
        print(len(week2a))
        print(len(week1b))
        print(len(week2b))

        '''defining the calcspear function and findng the spearman coefficient value'''


        def calcSpear(la, lb):
            return scipy.stats.spearmanr(la, lb)[0]


        r1a2a = calcSpear(week1a, week2a)
        r1a2b = calcSpear(week1a, week2b)
        r2a2b = calcSpear(week2a, week2b)
        if(math.isnan(r1a2a)):
            r1a2a = 0.0
        if(math.isnan(r1a2b)):
            r1a2b = 0.0
        if(math.isnan(r2a2b)):
            r2a2b = 0.0
        if(r1a2a == 1):
            r1a2a = 0.99
        if(r1a2b == 1):
            r1a2b = 0.99
        if(r2a2b == 1):
            r2a2b = 0.99
        N = len(week1a)

        '''calculating the calz value according to given function'''


        def calcZ(r1a2a, r1a2b, r2a2b, N):

            rm2 = ((r1a2a ** 2) + (r1a2b ** 2)) / 2
            f = (1 - r2a2b) / (2 * (1 - rm2))
            h = (1 - f * rm2) / (1 - rm2)

            z1a2a = 0.5 * (math.log10((1 + r1a2a) / (1 - r1a2a)))
            z1a2b = 0.5 * (math.log10((1 + r1a2b) / (1 - r1a2b)))

            z = (z1a2a - z1a2b) * ((N - 3) ** 0.5) / (2 * (1 - r2a2b) * h)

            return z




        z=calcZ(r1a2a,r1a2b,r2a2b,N)
        '''finding the p(z) value'''


        def calcP(z):
            p = 0.3275911
            a1 = 0.254829592
            a2 = -0.284496736
            a3 = 1.421413741
            a4 = -1.453152027
            a5 = 1.061405429

            sign = None
            if z < 0.01:
                sign = -1
            else:
                sign = 1

            x = abs(z) / (2 ** 0.5)
            t = 1 / (1 + p * x)
            erf = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * math.exp(-x * x)

            return 0.5 * (1 + sign * erf)




        print(calcP(z))
        sheetsecond.write(a,b,calcP(z))

        '''' saving the xls file'''

b1.save("p227sec.xls")
b1.save(TemporaryFile())







