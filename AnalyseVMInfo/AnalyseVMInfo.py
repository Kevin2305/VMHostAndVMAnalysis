#!/usr/bin/python
# coding: utf-8
import xlrd
import xlsxwriter
import GlobalVariables

def openExcel(filename):
    try:
        wb = xlrd.open_workbook(filename)
        return wb
    except Exception as e:
        print(e)

def getData(workbook_obj,sheetname,flag='row'):
    ws = workbook_obj.sheet_by_name(sheetname)
    data = []
    if flag == 'row':
        for i in range(0,ws.nrows):
            data.append(ws.row_values(i))
    elif flag == 'column':
        for i in range(0,ws.ncols):
            data.append(ws.col_values(i))
    else:
        print('flag error')
        return
    return data

def writeExcel(sheetname_obj,content,x_axis,y_axis):
    try:
        sheetname_obj.write(x_axis,y_axis,content)
    except Exception as e:
        print(e)

def drawChart(workbook_obj,sheetname_obj,sheet1_nrow,sheet2_nrom):
    newsheet_nrow = sheet1_nrow + sheet2_nrom
    main_chart = workbook_obj.add_chart({'type': 'column'})
    sec_chart = workbook_obj.add_chart({'type': 'line'})
    main_chart.add_series({'categories': '=%s!A2:A%s' % (GlobalVariables.DEST_SHEET_NAME,newsheet_nrow),
                           'values': '=%s!B2:B%s' % (GlobalVariables.DEST_SHEET_NAME,newsheet_nrow),
                           'name': 'Total vCPU(#)',
                           'fill': {'color': 'red'}})
    main_chart.add_series({'values': '=%s!C2:C%s' % (GlobalVariables.DEST_SHEET_NAME,newsheet_nrow),
                           'name': 'Provisioned vCPU(#)',
                           'fill': {'color': 'yellow'}})
    main_chart.add_series({'values':'=%s!D2:D%s' % (GlobalVariables.DEST_SHEET_NAME,newsheet_nrow),
                           'name':'Total Physical Memory(GB)'})
    main_chart.add_series({'values':'=%s!E2:E%s' % (GlobalVariables.DEST_SHEET_NAME,newsheet_nrow),
                           'name':'Provisioned Physical Memory(GB)'})
    sec_chart.add_series({'categories':'=%s!A2:A%s' % (GlobalVariables.DEST_SHEET_NAME,newsheet_nrow),
                          'values': '=%s!G2:G%s' % (GlobalVariables.DEST_SHEET_NAME,newsheet_nrow),
                          'marker': {'type': 'diamond',
                                     'size': 8,
                                     #'border': {'color': 'black'},
                                     'fill':   {'color': 'blue'}
                                     },
                          'name': 'Spare VM Count(#)',
                          'y2_axis': True,
                          'line': {'size': 3,
                                   'color': 'green',
                                   'none': True}
                         })
    main_chart.set_size({'width': GlobalVariables.MAIN_CHART_WIDTH, 'height': GlobalVariables.MAIN_CHART_HEIGHT})
    main_chart.set_title({'name': GlobalVariables.MAIN_CHART_TITLE})
    main_chart.set_x_axis({'name': GlobalVariables.MAIN_CHART_X_NAME})
    main_chart.set_y_axis({'name': GlobalVariables.MAIN_CHART_Y_NAME})

    main_chart.set_table({'show_keys': True})
    main_chart.set_legend({'position': 'none'})

    sec_chart.set_y2_axis({'name': GlobalVariables.SEC_CHART_Y2_NAME})

    main_chart.combine(sec_chart)
    sheetname_obj.insert_chart('A%d' % (newsheet_nrow+2),main_chart)

def main():
    # open source xlsx
    src_wb = openExcel(GlobalVariables.FILEPATH + GlobalVariables.SRC_FILE)

    # Get src data
    #sheet3 = getData(src_wb,GlobalVariables.SRC_SHEET_NAME, 'row')
    sheet1 = getData(src_wb,GlobalVariables.VMHOSTS, 'column')
    sheet2 = getData(src_wb,GlobalVariables.VMS, 'row')
    
    
    with xlsxwriter.Workbook(GlobalVariables.FILEPATH + GlobalVariables.DEST_FILE) as dest_wb:
        dest_sh_stat = dest_wb.add_worksheet(GlobalVariables.DEST_SHEET_NAME)
        sheet1_nrow = vmhostsStatistic(dest_sh_stat,sheet1,sheet2)
        #sheet2_nrow = clusterStatistic(dest_sh_stat,sheet1,sheet2)
    '''
        for x in range(len(sheet3)):
            for y in range(len(sheet3[x])):
                writeExcel(dest_sh_stat,sheet3[x][y],x,y)
            vmAnalysis(dest_wb,dest_sh_stat,sheet3,sheet1)
    '''

def vmhostsStatistic(sheetname,vmhosts,vms):    
    title = ['VM Host Name',
             'Total vCPU(#)',
             'Provisioned vCPU(#)',
             'Total Physical Memory(GB)',
             'Provisioned Physical Memory(GB)',
             'Current VM Count(#)',
             'Spare VM Count(#)']
    sheetname.write_row('A1', title)
    hostnames = list(set(vmhosts[0]) - set([vmhosts[0][0]]))
    dict_hosts = {}
    for cname in hostnames:
        dict_hosts['%s' % cname]={'totalvcpu': 0,
                                  'totalvmem': 0,
                                  'totalusedvcpu': 0,
                                  'totalusedvmem': 0,
                                  'totalvmnum': 0,
                                  'totalsparevmnum': 0}
    for i in range(1,len(vms)):
        host = vms[i][5]
        for j in range(1,len(vmhosts[0])):
            if vmhosts[0][j] == host:
                if vmhosts[6][j] == 1 or vmhosts[6][j] == True:
                    dict_hosts[host]['totalvcpu'] = vmhosts[1][j] * 2 
                else:
                    dict_hosts[host]['totalvcpu'] = vmhosts[1][j]
                dict_hosts[host]['totalvmem'] = vmhosts[2][j]
        dict_hosts[host]['totalusedvcpu'] = dict_hosts[host]['totalusedvcpu'] + vms[i][1]
        dict_hosts[host]['totalusedvmem'] = dict_hosts[host]['totalusedvmem'] + vms[i][2]
        dict_hosts[host]['totalvmnum'] = dict_hosts[host]['totalvmnum'] + 1
    for name in hostnames:
        if dict_hosts[name]['totalvmem'] <= dict_hosts[name]['totalusedvmem']:
            dict_hosts[name]['totalsparevmnum'] = 0
        else:
            dict_hosts[name]['totalsparevmnum'] = (dict_hosts[name]['totalvmem'] - dict_hosts[name]['totalusedvmem'])//GlobalVariables.STD_VM_MEM
    content = formatDict2List(dict_hosts)
    for i in range(len(content)):
        sheetname.write_row('A%s' % (i + 2), content[i])
    return len(hostnames)

def clusterStatistic(sheetname,vmhosts,vmstat):
    # cluster capacity analysis
    clusternames = list(set(vmhosts[8]) - set([vmhosts[8][0]]))
    dict_cls = {}
    for cname in clusternames:
        dict_cls['%s' % cname]={'totalvcpu': 0,
                                'totalvmem': 0,
                                'totalusedvcpu': 0,
                                'totalusedvmem': 0,
                                'totalvmnum': 0,
                                'totalsparevmnum': 0}
    for i in range(1,len(vmhosts[0])):
        cluster = vmhosts[8][i]
        vmhost = vmhosts[0][i]
        for j in range(1,len(vmstat)):
            if vmhost == vmstat[j][0]:
                dict_cls[cluster]['totalvcpu'] = dict_cls[cluster]['totalvcpu'] + vmstat[j][1]
                dict_cls[cluster]['totalusedvcpu'] = dict_cls[cluster]['totalusedvcpu'] + vmstat[j][2]
                dict_cls[cluster]['totalvmem'] = dict_cls[cluster]['totalvmem'] + vmstat[j][3]
                dict_cls[cluster]['totalusedvmem'] = dict_cls[cluster]['totalusedvmem'] + vmstat[j][4]
                dict_cls[cluster]['totalvmnum'] = dict_cls[cluster]['totalvmnum'] + vmstat[j][5]
                dict_cls[cluster]['totalsparevmnum'] = dict_cls[cluster]['totalsparevmnum'] + vmstat[j][6]
    for i in range(len(clusternames)):
        content = []
        content.append(clusternames[i])
        content.append(dict_cls[clusternames[i]]['totalvcpu'])
        content.append(dict_cls[clusternames[i]]['totalusedvcpu'])
        content.append(dict_cls[clusternames[i]]['totalvmem'])
        content.append(dict_cls[clusternames[i]]['totalusedvmem'])
        content.append(dict_cls[clusternames[i]]['totalvmnum'])
        content.append(dict_cls[clusternames[i]]['totalsparevmnum'])
        sheetname.write_row('A%s' % (len(vmstat) + 1 + i), content)
    return len(clusternames)

def formatDict2List(dict_data):
    keys = list(dict_data.keys())
    result = []
    for i in range(len(keys)):
        content = []
        content.append(keys[i])
        content.append(dict_data[keys[i]]['totalvcpu'])
        content.append(dict_data[keys[i]]['totalusedvcpu'])
        content.append(dict_data[keys[i]]['totalvmem'])
        content.append(dict_data[keys[i]]['totalusedvmem'])
        content.append(dict_data[keys[i]]['totalvmnum'])
        content.append(dict_data[keys[i]]['totalsparevmnum'])
        result.append(content)
    return result

def get_VMHost_vCPU_vMEM(vmhosts, hostname):
    data = {'totalvcpu': 0,
            'totalvmem': 0}
    for i in range(len(1,vmhosts[0])):
        if vmhosts[0][i] == hostname:
            if vmhosts[6][i] == True or vmhosts[6][i] == 'True' or vmhosts[6][i] ==1:
                data['totalvcpu'] = (vmhosts[1][i] * 2)
            else:
                data['totalvcpu'] = vmhosts[1][i]
            data['totalvmem'] = vmhosts[2][i]
            return data
    return 0


if __name__ == '__main__':
    main()