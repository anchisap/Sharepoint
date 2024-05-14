from Sharepoint import SharePoint
from openpyxl import Workbook

#get CL1_S1007_Absence_Request sharepoint list
CL1_S1007_Absence_Request = SharePoint().connect_to_list(ls_name='CL1_S1007_Absence_Request')


print(CL1_S1007_Absence_Request)