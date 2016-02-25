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

def analysisVM(workbook_obj,sheetname_obj,data):
    main_chart = workbook_obj.add_chart({'type':'column'})
    sec_chart = workbook_obj.add_chart({'type':'line'})
    main_chart.add_series({'categories':'=%s!A2:A%s' % (GlobalVariables.DEST_SHEET_NAME,len(data)),
                           'values':'=%s!B2:B%s' % (GlobalVariables.DEST_SHEET_NAME,len(data)),
                           'name':'%s' % (data[0][1]),
                           'fill': {'color': 'red'}})
    main_chart.add_series({'values':'=%s!C2:C%s' % (GlobalVariables.DEST_SHEET_NAME,len(data)),
                           'name':'%s' % (data[0][2]),
                           'fill': {'color': 'yellow'}})
    main_chart.add_series({'values':'=%s!D2:D%s' % (GlobalVariables.DEST_SHEET_NAME,len(data)),
                           'name':'%s' % (data[0][3])})
    main_chart.add_series({'values':'=%s!E2:E%s' % (GlobalVariables.DEST_SHEET_NAME,len(data)),
                           'name':'%s' % (data[0][4])})
    sec_chart.add_series({'categories':'=%s!A2:A%s' % (GlobalVariables.DEST_SHEET_NAME,len(data)),
                          'values': '=%s!G2:G%s' % (GlobalVariables.DEST_SHEET_NAME,len(data)),
                          'marker': {'type': 'diamond',
                                     'size': 8,
                                     #'border': {'color': 'black'},
                                     'fill':   {'color': 'blue'}
                                     },
                          'name': '%s' % (data[0][6]),
                          'y2_axis': True,
                          'line': {'size': 3,
                                   'color': 'green',
                                   'none': True}
                         })
    main_chart.set_size({'width': GlobalVariables.MAIN_CHART_WIDTH, 'height': GlobalVariables.MAIN_CHART_HEIGHT})
    main_chart.set_title({'name': GlobalVariables.MAIN_CHART_TITLE})
    main_chart.set_x_axis({'name': GlobalVariables.MAIN_CHART_X_NAME})
    main_chart.set_y_axis({'name': GlobalVariables.MAIN_CHART_Y_NAME})
    sec_chart.set_y2_axis({'name': GlobalVariables.SEC_CHART_Y2_NAME})

    main_chart.combine(sec_chart)
    sheetname_obj.insert_chart('A%d' % (len(data)+2),main_chart)

def main():
    # open source xlsx
    src_wb = openExcel(GlobalVariables.FILEPATH + GlobalVariables.SRC_FILE)

    # Get src data
    sheet3 = getData(src_wb,GlobalVariables.SRC_SHEET_NAME, 'row')
    sheet1 = getData(src_wb,GlobalVariables.VMHOST, 'column')
    clusterCapacity(sheet1,sheet3)
    '''
    with xlsxwriter.Workbook(GlobalVariables.FILEPATH + GlobalVariables.DEST_FILE) as dest_wb:
        dest_sh_stat = dest_wb.add_worksheet(GlobalVariables.DEST_SHEET_NAME)
        for x in range(len(sheet3)):
            for y in range(len(sheet3[x])):
                writeExcel(dest_sh_stat,sheet3[x][y],x,y)
            analysisVM(dest_wb,dest_sh_stat,sheet3)
    '''

def clusterCapacity(vmhosts,vmstat):
    # cluster capacity
    clusternames = list(set(vmhosts[8]) - set([vmhosts[8][0]]))
    for cname in clusternames:
        dict_cls['%s' % cname]={'totalvcpu': 0,
                                'totalvmem': 0,
                                'totalusedvcpu': 0,
                                'totalusedvmem': 0}
    for i in range(1,len(vmhosts[0])):
        cluster = vmhosts[8][i]
        vmhost = vmhosts[1][i]
        for j in range(1,len(vmstat)):
            if vmhost == vmstat[j][0]:




 


if __name__ == '__main__':
    main()