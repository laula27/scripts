import sys              # Get program arguments
import os               # Enter system commands
import pandas as pd
import numpy as np
import xlsxwriter       # Write and format xlsx files
import csv
import re
import subprocess

def extract_ops_data(tww):
    data = ['PGSRT','HSRT']
    configuration = ['X8','X16']
    fab = ['FAB11','FAB16']
    for d in data:
        for f in fab:
            for c in configuration:
                onc = 'cp /vol/pye/MTI/OPS/Z41C/'+str(tww)+'/DATA/'+d+'/'+f+'/'+c+'/SDP/FIRST_PASS/SUMMARIES z41c_'+f+'_'+d+'_'+c+'_SDP_ww'+str(tww)+'.txt'
                os.system(onc)
                #subprocess.run('cp /vol/pye/MTI/OPS/Z41C/'+str(tww)+'/DATA/'+d+'/'+f+'/'+c+'/SDP/FIRST_PASS/SUMMARIES z41c_'+f+'_'+d+'_'+c+'_SDP_ww'+str(tww)+'.txt',shell = True)
                #tsum =  'tsums @z41c_'+f+'_'+d+'_'+c+'_SDP_ww'+str(tww)+'.txt -format=MAJOR_PROBE_PROG_REV,' 
                #command = 'tsums @z41c_'+f+'_'+d+'_'+c+'_SDP_ww'+str(tww)+'.txt -format=MAJOR_PROBE_PROG_REV, "sum(uin)", >| '+f+'_'+c+'_'+d+'_mppr.txt'
                #output  = subprocess.Popen(command, universal_newlines=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
                #command = ['tsums @z41c_'+f+'_'+d+'_'+c+'_SDP_ww'+str(tww)+'.txt -format=RETICLE_WAVE_ID,", "sum(uin)", >| '+f+'_'+c+'_'+d+'_wave.txt']
                #output  = subprocess.Popen(command, universal_newlines=True, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
                tsum =  'tsums @z41c_'+f+'_'+d+'_'+c+'_SDP_ww'+str(tww)+'.txt -format=MAJOR_PROBE_PROG_REV,' 
                txt = ' | sort > '+f+'_'+c+'_'+d+'_mppr.txt'
                subprocess.run(tsum + "'sum(uin)'" + txt,shell = True)
                tsum =  'tsums @z41c_'+f+'_'+d+'_'+c+'_SDP_ww'+str(tww)+'.txt -format=RETICLE_WAVE_ID,' 
                txt = ' | sort > '+f+'_'+c+'_'+d+'_wave.txt'
                subprocess.run(tsum + "'sum(uin)'" + txt,shell = True)
                excel_ops_data(''+f+'_'+c+'_'+d+'_mppr.txt',''+f+'_'+c+'_'+d+'_wave.txt',f,c,d,tww)
               
                


def excel_ops_data(file1,file2,fab,c,d,tww): 
    mppr_file = open(file1, "r")
    mpprtxt = mppr_file.read()
    mpprgrp = re.findall(r'\s*(\d+)\s+(\d+)',mpprtxt)
    wave_file = open(file2, "r")
    wavetxt = wave_file.read()
    wavegrp = re.findall(r'\s*(\w+)\s+(\d+)',wavetxt)
    total = 0
    for mtuple in mpprgrp:
        x = int(mtuple[1])
        total = total + x
    tuple_new_mppr = [(x, y, '{0:.2f}'.format(int(y) * 100 /total)) for x, y in mpprgrp]
    total = 0
    for wavetuple in wavegrp:
        print(wavetuple)
    for wavetuple in wavegrp:
        x = int(wavetuple[1])
        total = total + x
    tuple_new_wave = [(x, y, '{0:.2f}'.format(int(y) * 100 /total)) for x, y in wavegrp]
    mppr_file.close()
    wave_file.close()
    save_to_txt(tuple_new_mppr,tuple_new_wave,fab,c,d,tww)

def save_to_txt(tuple_new_mppr, tuple_new_wave,fab,c,d,tww):
    with open('temp.txt', 'w') as f:
        for t in tuple_new_mppr:
            f.write(','.join(str(s) for s in t) + ",\n")
    with open('temp2.txt', 'w') as f:
        for t in tuple_new_wave:
            f.write(','.join(str(s) for s in t) + ',\n')
    cmd ='paste temp.txt temp2.txt > temp3.txt'
    os.system(cmd)
    save_to_excel(fab,c,d,tww)

def save_to_excel(fab,c,d,tww):
    with open('temp3.txt', 'r') as in_file:
        stripped = (line.strip() for line in in_file)
        lines = (line.split(",") for line in stripped if line)
        print(lines)
        with open(''+tww+'.csv', 'a',encoding='UTF-8') as out_file:
            writer = csv.writer(out_file)
            writer.writerow((fab,d,c))
            writer.writerow(('mppr','amount','percentage','wave','amount','percentage'))
            writer.writerows(lines)
    cmd ='rm *.txt'
    os.system(cmd)
def main():  
    tww_passed = False
    if ("help" in str(sys.argv)):
        print("run python3")
    for argument in sys.argv:
        if "tww" in argument: 
            tww_passed = True
    if not tww_passed:
        tww = input("tww(ex:202131): ")
    if os.path.exists(''+tww+'.csv'):
        os.remove(''+tww+'.csv')
    extract_ops_data(tww)

if __name__=="__main__":
    main()