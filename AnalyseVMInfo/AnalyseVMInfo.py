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
    newsheet_nrow = sheet1_nrow + sheet2_nrom + 1
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
                          'name': 'Available VM Count(#)',
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
    sheetname_obj.insert_chart('A%d' % (newsheet_nrow+3),main_chart)


def vmhostsStatistic(sheetname,vmhosts,vms):    
    title = ['VM Host Name',
             'Total vCPU(#)',
             'Provisioned vCPU(#)',
             'Total Physical Memory(GB)',
             'Provisioned Physical Memory(GB)',
             'Current VM Count(#)',
             'Available VM Count(#)']
    sheetname.write_row('A1', title)
    hostnames = list(set(vmhosts[0]) - set([vmhosts[0][0]]))
    dict_hosts = {}
    for cname in hostnames:
        dict_hosts['%s' % cname] = {'totalvcpu': 0,
                                    'totalvmem': 0,
                                    'totalusedvcpu': 0,
                                    'totalusedvmem': 0,
                                    'totalvmnum': 0,
                                    'totalsparevmnum': 0}
    for name in hostnames:
        vmhost_info = get_VMHost_vCPU_vMEM(vmhosts,name)
        vmhost_used_info = get_VMHost_Used_vCPU_vMEM(vms,name, (GlobalVariables.VMHOST_COL_IN_VMS_SHEET -1))
        dict_hosts[name]['totalvcpu'] = vmhost_info['totalvcpu']
        dict_hosts[name]['totalvmem'] = vmhost_info['totalvmem']
        dict_hosts[name]['totalusedvcpu'] = vmhost_used_info['totalusedvcpu']
        dict_hosts[name]['totalusedvmem'] = vmhost_used_info['totalusedvmem']
        dict_hosts[name]['totalvmnum'] = vmhost_used_info['totalvmnum']
        if dict_hosts[name]['totalvmem'] <= dict_hosts[name]['totalusedvmem']:
            dict_hosts[name]['totalsparevmnum'] = 0
        else:
            dict_hosts[name]['totalsparevmnum'] = (dict_hosts[name]['totalvmem'] - dict_hosts[name]['totalusedvmem'])//GlobalVariables.STD_VM_MEM
    content = formatDict2List(dict_hosts)
    for i in range(len(content)):
        sheetname.write_row('A%s' % (i + 2), content[i])
    return len(hostnames)


def clusterStatistic(sheetname,vmhosts,vms):
    # cluster capacity analysis
    clusternames = list(set(vmhosts[(GlobalVariables.VMCLUSTER_COL_IN_VMHOSTS_SHEET - 1)]) - set([vmhosts[(GlobalVariables.VMCLUSTER_COL_IN_VMHOSTS_SHEET - 1)][0]]))
    dict_cls = {}
    for cname in clusternames:
        dict_cls['%s' % cname]={'totalvcpu': 0,
                                'totalvmem': 0,
                                'totalusedvcpu': 0,
                                'totalusedvmem': 0,
                                'totalvmnum': 0,
                                'totalsparevmnum': 0}
    for name in clusternames:
        cluster_info = get_VMHost_vCPU_vMEM(vmhosts, name, (GlobalVariables.VMCLUSTER_COL_IN_VMHOSTS_SHEET - 1))
        cluster_used_info = get_VMHost_Used_vCPU_vMEM(vms, name , (GlobalVariables.VMCLUSTER_COL_IN_VMS_SHEET - 1))
        dict_cls[name]['totalvcpu'] = cluster_info['totalvcpu']
        dict_cls[name]['totalvmem'] = cluster_info['totalvmem']
        dict_cls[name]['totalusedvcpu'] = cluster_used_info['totalusedvcpu']
        dict_cls[name]['totalusedvmem'] = cluster_used_info['totalusedvmem']
        dict_cls[name]['totalvmnum'] = cluster_used_info['totalvmnum']
        if dict_cls[name]['totalvmem'] <= dict_cls[name]['totalusedvmem']:
            dict_cls[name]['totalsparevmnum'] = 0
        else:
            dict_cls[name]['totalsparevmnum'] = (dict_cls[name]['totalvmem'] - dict_cls[name]['totalusedvmem'])//GlobalVariables.STD_VM_MEM
    content = formatDict2List(dict_cls)
    for i in range(len(content)):
        sheetname.write_row('A%s' % (len(vmhosts[0]) + 1 + i), content[i])
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


def get_VMHost_vCPU_vMEM(vmhosts_sheet, hostname, *args):
    data = {'totalvcpu': 0,
            'totalvmem': 0}
    if len(args) <= 0:
        for i in range(1, len(vmhosts_sheet[0])):
            if vmhosts_sheet[0][i] == hostname:
                if vmhosts_sheet[6][i] == True or vmhosts_sheet[6][i] == 'True' or vmhosts_sheet[6][i] ==1:
                    data['totalvcpu'] = (vmhosts_sheet[1][i] * 2)
                else:
                    data['totalvcpu'] = vmhosts_sheet[1][i]
                data['totalvmem'] = vmhosts_sheet[2][i]
                return data
    else:
        for i in range(1, len(vmhosts_sheet[args[0]])):
            if vmhosts_sheet[args[0]][i] == hostname:
                 vmhost_info = get_VMHost_vCPU_vMEM(vmhosts_sheet, vmhosts_sheet[0][i])
                 data['totalvcpu'] = data['totalvcpu'] + vmhost_info['totalvcpu']
                 data['totalvmem'] = data['totalvmem'] + vmhost_info['totalvmem']
        return data
    return


def get_VMHost_Used_vCPU_vMEM(vms_sheet, hostname, col):
    # get vm host or vm cluster used vcpu and vmem
    # col is the column id of vm host name or cluster name
    data = {'totalusedvcpu': 0,
            'totalusedvmem': 0,
            'totalvmnum': 0 }
    result = []
    for i in range(1, len(vms_sheet)):
        if vms_sheet[i][col] == hostname:
            data['totalusedvcpu'] = data['totalusedvcpu'] + vms_sheet[i][1]
            data['totalusedvmem'] = data['totalusedvmem'] + vms_sheet[i][2]
            data['totalvmnum'] = data['totalvmnum'] + 1
    return data



def main():
    # open source xlsx
    src_wb = openExcel(GlobalVariables.FILEPATH + GlobalVariables.SRC_FILE)

    # Get src data
    #sheet3 = getData(src_wb,GlobalVariables.SRC_SHEET_NAME, 'row')
    sheet1 = getData(src_wb,GlobalVariables.VMHOSTS, 'column')
    sheet2 = getData(src_wb,GlobalVariables.VMS, 'row')
    
    
    with xlsxwriter.Workbook(GlobalVariables.FILEPATH + GlobalVariables.DEST_FILE) as dest_wb:
        dest_sheet = dest_wb.add_worksheet(GlobalVariables.DEST_SHEET_NAME)
        sheet1_nrow = vmhostsStatistic(dest_sheet,sheet1,sheet2)
        sheet2_nrow = clusterStatistic(dest_sheet,sheet1,sheet2)
        drawChart(dest_wb,dest_sheet,sheet1_nrow,sheet2_nrow)


if __name__ == '__main__':
    main()