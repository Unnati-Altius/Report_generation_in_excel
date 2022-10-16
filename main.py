import xlsxwriter
from Report.Report import RfpReport
from Report.PricesReport import PReport

# create file(workbook) and worksheet
workbook = xlsxwriter.Workbook('Report.xlsx')
worksheet1 = workbook.add_worksheet()
worksheet2 = workbook.add_worksheet()

# declare data
company = ["HITACHI VANTARA", "COFORGE", "INFOSYS"]

services = ["EUC/Desktop Operations", "Messaging, Collaboration, Sharepoint and 0365 Services", "Cloud operations Support",
"LAN, WAN Support","Data Center Operations", "DB Support", "ITSM Support", "Governance", "Additional EUC Scope"]

mapping = {"HITACHI VANTARA" : [["EUC/Desktop Operations", 24],
["Messaging, Collaboration, Sharepoint and 0365 Services", 30], ["Cloud operations Support", 20], ["LAN, WAN Support", 34], 
["Data Center Operations", 15], ["DB Support", 25], [ "ITSM Support", 40], ["Governance", 35]], "COFORGE": [["EUC/Desktop Operations", 27],
["Messaging, Collaboration, Sharepoint and 0365 Services", 50], ["Cloud operations Support", 34], ["LAN, WAN Support", 12], ["Data Center Operations", 27],
["DB Support", 42], ["Governance", 24], ["Additional EUC Scope", 23]], "INFOSYS": [["EUC/Desktop Operations", 30], 
["Cloud operations Support", 16], ["LAN, WAN Support", 30], 
["Data Center Operations", 18], ["DB Support", 40], [ "ITSM Support", 28], ["Governance", 17]]}


if __name__ == "__main__":
    '''
    This is for sheet1 representing yes/no
    '''
    worksheet1.write("A1", "Tower")
    worksheet1.write(1, 0, "At a glance")
    worksheet1.set_column(0, len(company), 45)

    rfp_report_instance = RfpReport(worksheet1, company, services, mapping)
    rfp_report_instance.add_companies()
    rfp_report_instance.add_services()

    '''
    This is for sheet2 representing total cost a vendor is taking
    '''
    worksheet2.write("A1", "Tower")
    worksheet2.write(1, 0, "At a glance")
    worksheet2.set_column(0, len(company), 45)

    p_report_instance = PReport(worksheet2, company, services, mapping)
    p_report_instance.add_companies()
    p_report_instance.add_services()
    
    workbook.close()
    
