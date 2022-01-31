# -*- coding: utf-8 -*-

import sys
import pickle
import calendar
from pyzabbix import ZabbixAPI
from datetime import datetime, timedelta
import time
import subprocess
from xlrd import open_workbook
from xlutils.copy import copy

months = ["Unknown",
          "январь",
          "февраль",
          "март",
          "апрель",
          "май",
          "июнь",
          "июль",
          "август",
          "сентябрь",
          "октябрь",
          "ноябрь",
          "декабрь"]

def events_gr_4h(events):
    result = []
    for event in events:
        if (event['r_eventid'] != '0'):
            r_event = 0
            try:
                r_event = [x for x in events if x['eventid'] == event['r_eventid']][0]
            except Exception as e:
                print(e)
            if (r_event and (float(r_event['clock']) - float(event['clock']) > 14400)):
                event['r_clock'] = r_event['clock']
                result.append(event)
    return result

def events_description(events1, events2):
    print('----------------------------------------------------------------------------------')
    description = ''
    for e1 in events1:
        host = e1['hosts'][0]['host']
        eventid = e1['eventid']
        start_time = datetime.fromtimestamp(float(e1['clock'])).strftime("%Y.%m.%d %H:%M")
        end_time = datetime.fromtimestamp(float(e1['r_clock'])).strftime("%Y.%m.%d %H:%M")
        cross = False
        clock_7200 = float(e1['clock']) + 7200
        for e2 in events2:
            if ((float(e2['clock']) < clock_7200) and (clock_7200 < float(e2['r_clock']))):
                cross = True
        if cross:
            description = description + '\n' + start_time + ' - ' + end_time + ' Откл. эл.'
        else:
            description = description + '\n' + start_time + ' - ' + end_time
        print('event', host, eventid, start_time, end_time)
    return description

ifindexes = {}
file = open('ifindexes.txt', 'rb')
ifindexes = pickle.load(file)
file.close()
zapi=ZabbixAPI('http://localhost/zabbix', user='Admin', password='******************************')
book = copy(open_workbook('./tmp.xls', formatting_info=True, encoding_override="cp1251"))
sheet = book.get_sheet(0)

date_till = datetime.today().replace(day=1) - timedelta(days=1)
for i in range(int(sys.argv[1]) - 1):
    date_till = date_till.replace(day=1) - timedelta(days=1)
date_from = date_till.replace(day=1)
date_till = datetime(date_till.year, date_till.month, date_till.day, 23, 59, 59)
date_from = datetime(date_from.year, date_from.month, date_from.day, 0, 0, 1)
time_from = int(time.mktime(date_from.timetuple()))
time_till = int(time.mktime(date_till.timetuple()))
header = 'ОПФР по Алтайскому краю за ' + months[int(date_from.month)]  + ' ' + str(date_from.year) + ' года'
sheet.write(3, 0, header.decode('utf-8'))
print('date_from: ', date_from, 'date_till: ', date_till, 'time from: ', time_from, ' time till: ', time_till)


events = zapi.event.get(time_from=time_from, time_till=time_till, select_acknowledges='extend', selectHosts='extend', filter={'name':'ICMP ERR'})
lines = open('data', 'r').readlines()
i = 7
j = 0
for line in lines:
    try:
        j = j + 1
        i = i + 2
        print('##################################################################################################################')
        data = line.split(';')
        lo_meg = data[5].strip()
        lo_rtk = data[5].strip()
        meg = data[3].strip()
        rtk = data[4].strip()
        name = data[1]
        address = data[2]
        if meg == '10.212.32.1':
            lo_meg = '10.32.254.254'
            lo_rtk = '10.32.255.254'
        rtk_speed = int(data[6]) * 1024
        meg_speed = int(data[7]) * 1024
        print(name)
        print(lo_rtk,lo_meg, rtk, meg)

        items_rtk = zapi.host.get(time_from=time_from, time_till=time_till, selectItems='extend', output='extend', filter={'host':lo_rtk})
        items_meg = zapi.host.get(time_from=time_from, time_till=time_till, selectItems='extend', output='extend', filter={'host':lo_meg})
        try:
            rtk_snmp_ifindex,stderr = subprocess.Popen(['snmpwalk', '-Oqv', '-v', '2c', '-c', '******************************', lo_rtk, 'IP-MIB::ipAdEntIfIndex.' + rtk], stdout=subprocess.PIPE, stderr=subprocess.STDOUT).communicate()
            rtk_snmp_ifindex = str(int(rtk_snmp_ifindex.strip()))
            meg_snmp_ifindex,stderr = subprocess.Popen(['snmpwalk', '-Oqv', '-v', '2c', '-c', '******************************', lo_meg, 'IP-MIB::ipAdEntIfIndex.' + meg], stdout=subprocess.PIPE, stderr=subprocess.STDOUT).communicate()
            meg_snmp_ifindex = str(int(meg_snmp_ifindex.strip()))
            ifindexes[lo_rtk] = {}
            ifindexes[lo_meg] = {}
            ifindexes[lo_rtk]['rtk_snmp_ifindex'] = rtk_snmp_ifindex
            ifindexes[lo_meg]['meg_snmp_ifindex'] = meg_snmp_ifindex
            file = open('ifindexes.txt', 'wb')
            pickle.dump(ifindexes, file)
            file.close()
        except Exception as e:
            print(e)
            rtk_snmp_ifindex = ifindexes[lo_rtk]['rtk_snmp_ifindex']
            meg_snmp_ifindex = ifindexes[lo_meg]['meg_snmp_ifindex']

        print('rtk_snmp_ifindex: ', rtk_snmp_ifindex)
        print('meg_snmp_ifindex:', meg_snmp_ifindex)

        rtk_received_itemid = int([item for item in items_rtk[0]['items'] if '1.3.6.1.2.1.31.1.1.1.6.' + rtk_snmp_ifindex == item['snmp_oid']][0]['itemid'])
        rtk_sent_itemid = [item for item in items_rtk[0]['items'] if '1.3.6.1.2.1.31.1.1.1.10.' + rtk_snmp_ifindex == item['snmp_oid']][0]['itemid']
        meg_received_itemid = [item for item in items_meg[0]['items'] if '1.3.6.1.2.1.31.1.1.1.6.' + meg_snmp_ifindex == item['snmp_oid']][0]['itemid']
        meg_sent_itemid = [item for item in items_meg[0]['items'] if '1.3.6.1.2.1.31.1.1.1.10.' + meg_snmp_ifindex == item['snmp_oid']][0]['itemid']
        print('rtk_received_itemid: ', rtk_received_itemid)
        print('rtk_sent_itemid: ', rtk_sent_itemid)
        print('meg_received_itemid: ', meg_received_itemid)
        print('meg_sent_itemid: ', meg_sent_itemid)

        rtk_received_history = zapi.history.get(itemids=rtk_received_itemid, time_from=time_from, time_till=time_till, output='extend')
        rtk_sent_history = zapi.history.get(itemids=rtk_sent_itemid, time_from=time_from, time_till=time_till, output='extend')
        meg_received_history = zapi.history.get(itemids=meg_received_itemid, time_from=time_from, time_till=time_till, output='extend')
        meg_sent_history = zapi.history.get(itemids=meg_sent_itemid, time_from=time_from, time_till=time_till, output='extend')
        avg_rtk_received_speed = sum([int(item['value']) for item in rtk_received_history ])/len(rtk_received_history)
        avg_rtk_sent_speed = sum([int(item['value']) for item in rtk_sent_history ])/len(rtk_sent_history)
        avg_meg_received_speed = sum([int(item['value']) for item in meg_received_history ])/len(meg_received_history)
        avg_meg_sent_speed = sum([int(item['value']) for item in meg_sent_history ])/len(meg_sent_history)
        print('avg_rtk_received_speed: ', avg_rtk_received_speed)
        print('avg_rtk_sent_speed: ', avg_rtk_sent_speed)
        print('avg_meg_received_speed: ', avg_meg_received_speed)
        print('avg_meg_sent_speed: ', avg_meg_sent_speed)
    
        #rtk_transmited = (avg_rtk_received_speed+avg_rtk_sent_speed)*(time_till-time_from)/(1024*1024*8)
        rtk_transmited = avg_rtk_sent_speed*(time_till-time_from)/(1024*1024*8)
        meg_transmited = avg_meg_sent_speed*(time_till-time_from)/(1024*1024*8)
        print('rtk_transmited: ', rtk_transmited)
        print('meg_transmited: ', meg_transmited)
    
        events_rtk = events_gr_4h([e for e in events if rtk == e['hosts'][0]['host']])
        events_meg = events_gr_4h([e for e in events if meg == e['hosts'][0]['host']])
        events_rtk_duration = sum([float(e['r_clock']) - float(e['clock']) for e in events_rtk])
        events_meg_duration = sum([float(e['r_clock']) - float(e['clock']) for e in events_meg])
        rtk_events_count = len(events_rtk)
        meg_events_count = len(events_meg)
        rtk_description = events_description(events_rtk, events_meg)
        meg_description = events_description(events_meg, events_rtk)

        rtk_duration = str(int(events_rtk_duration / 3600)) + ':' + str(int(events_rtk_duration % 3600 / 60)) + ':' + str(int(events_rtk_duration % 3600 % 60))
        meg_duration = str(int(events_meg_duration / 3600)) + ':' + str(int(events_meg_duration % 3600 / 60)) + ':' + str(int(events_meg_duration % 3600 % 60))
        print(rtk_duration)
        print(meg_duration)
        print(rtk_description)
        print(meg_description)

        sheet.write(i, 0, j)
        sheet.write(i, 1, name.decode('utf-8'))
        sheet.write(i, 2, address.decode('utf-8'))
        sheet.write(i, 3, rtk_speed)
        sheet.write(i + 1, 3, meg_speed)
        sheet.write(i, 4, rtk_transmited)
        sheet.write(i + 1, 4, meg_transmited)
        if rtk_events_count:
            sheet.write(i, 7, rtk_events_count)
        if meg_events_count:
            sheet.write(i + 1, 7, meg_events_count)
        if rtk_events_count:
            sheet.write(i, 8, rtk_duration)
        if meg_events_count:
            sheet.write(i + 1, 8, meg_duration)
        sheet.write(i, 9, rtk_description.decode('utf-8'))
        sheet.write(i + 1, 9, meg_description.decode('utf-8'))

        book.save('/var/www/html/' + str(date_from.year) + '_' + str(date_from.month) + '.xls')
    except Exception as e:
        print(e)
