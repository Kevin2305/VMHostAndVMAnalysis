#!/usr/bin/python
# coding: utf-8

FILEPATH = 'C:\\Users\\admin\\Desktop\\'
SRC_FILE = 'uatenv.xlsx'
DEST_FILE = 'uatenv_dest.xlsx'
SRC_SHEET_NAME = 'Statistics'
DEST_SHEET_NAME = 'Statistics'
MAIN_CHART_X_NAME =  '宿主机名称' # 'VM Host Name'
MAIN_CHART_Y_NAME = '宿主机当前资源统计' # 'VM Host Resource Statistics'
SEC_CHART_Y2_NAME =  '可用虚拟机数量(#)' # 'Spare VM Count(#)'
MAIN_CHART_TITLE = '虚拟机可用数量 (STD: 2 vCPU/4G vMEM/60G HDD)'
MAIN_CHART_WIDTH = 1200
MAIN_CHART_HEIGHT = 580

VMHOSTS = 'VMHosts'
VMS = 'VMs'
STD_VM_MEM = 4
PIE_SHEET_NAME = 'Cluster Statistics'

VMHOST_COL_IN_VMS_SHEET = 6
VMCLUSTER_COL_IN_VMS_SHEET = 7
VMCLUSTER_COL_IN_VMHOSTS_SHEET = 9