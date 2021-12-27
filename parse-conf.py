import xlwt
import json
from xlwt import Workbook
from pathlib import Path
import re
import tkinter as tk
from tkinter import filedialog

def write_cell (sheet, row, col, str):
    try:
       sheet.write(row, col, str)
    except Exception as ex:     #Inteto de pisar celda
       row += 1
       sheet.write(row, col, str)
    return row

root = tk.Tk()
root.withdraw()

path = filedialog.askopenfilename()

print ("File: {}".format(path))
# ACI snapshot config file JSON
try:
    with open(path) as f:
        data = json.load(f)
except UnicodeDecodeError as uni:
    print ("USE A JSON FILE!!! Closing...")
    exit()

wb = Workbook()

#Pestaña Tenant policies information page
#wb = wb.add_sheet("Sheet 1", cell_overwrite_ok=True) #Permitir pisar celdas
#sheet3._cell_overwrite_ok = True/False permite modificar en el momento

for tn in data["polUni"]["children"]:           #Tenant Policies
    if "fvTenant" in tn and tn["fvTenant"]["attributes"]["name"]!= "common" and tn["fvTenant"]["attributes"]["name"]!= "infra":
        print("tn: {}".format(tn["fvTenant"]["attributes"]["name"]))
        sheet = wb.add_sheet(tn["fvTenant"]["attributes"]["name"])
        row = 0
        sheet.write(row, 0, "App Profile")
        sheet.write(row, 1, "EPG")
        sheet.write(row, 2, "BD")
        sheet.write(row, 3, "Subnets")
        sheet.write(row, 4, "VRF")
        sheet.write(row, 5, "Consume")
        sheet.write(row, 6, "Provide")
        sheet.write(row, 7, "Domains")
        sheet.write(row, 8, "Associations")
        sheet.write(row, 9, "Encap")
        for ap in tn["fvTenant"]["children"]:                                                           #Lista de dicts Tenants
            if "fvAp" in ap:
                print(" ap: {}".format(ap["fvAp"]["attributes"]["name"]))
                row += 1
                sheet.write(row, 0, ap["fvAp"]["attributes"]["name"])                    
                for epg in ap["fvAp"]["children"]:                                                      #Lista de dicts AP
                    if "fvAEPg" in epg:
                        print("     epg: {}".format(epg["fvAEPg"]["attributes"]["name"]))
                        row = write_cell(sheet, row, 1, epg["fvAEPg"]["attributes"]["name"])
                        rows = [row] * 8
                        for e in epg["fvAEPg"]["children"]:                                            #Lista de dicts EPGs
                            if "fvRsBd" in e:
                                print("     bd: {}".format(e["fvRsBd"]["attributes"]["tnFvBDName"]))
                                rows[0] = write_cell(sheet, rows[0], 2, e["fvRsBd"]["attributes"]["tnFvBDName"])
                                for bd in tn["fvTenant"]["children"]:
                                    if "fvBD" in bd and bd["fvBD"]["attributes"]["name"]==e["fvRsBd"]["attributes"]["tnFvBDName"]:
                                        for subnet in bd["fvBD"]["children"]: 
                                            if "fvSubnet" in subnet:   
                                                rows[1] = write_cell(sheet, rows[1], 3, subnet["fvSubnet"]["attributes"]["ip"])
                                            if "fvRsCtx" in subnet:
                                                rows[2] = write_cell(sheet, rows[2], 4, subnet["fvRsCtx"]["attributes"]["tnFvCtxName"])
                            if "fvRsDomAtt" in e:
                                print("     Dominio: {}".format(e["fvRsDomAtt"]["attributes"]["tDn"]))
                                rows[3] = write_cell(sheet, rows[3], 7, e["fvRsDomAtt"]["attributes"]["tDn"][4:])  #TODO Considerar todos los tipos de domino
                            if "fvRsPathAtt" in e:
                                print("     Association: {} | {}".format(e["fvRsPathAtt"]["attributes"]["tDn"],e["fvRsPathAtt"]["attributes"]["encap"]))
                                rows[4] = write_cell(sheet, rows[4], 8, e["fvRsPathAtt"]["attributes"]["tDn"][e["fvRsPathAtt"]["attributes"]["tDn"].find("[")+1:-1])
                                write_cell(sheet, rows[5], 9, e["fvRsPathAtt"]["attributes"]["encap"])
                            if "fvRsProv" in e:
                                print("     Provide: {}".format(e["fvRsProv"]["attributes"]["tnVzBrCPName"])) 
                                rows[6] = write_cell(sheet, rows[6], 6, e["fvRsProv"]["attributes"]["tnVzBrCPName"])   
                            if "fvRsCons" in e:
                                print("     Consume: {}".format(e["fvRsCons"]["attributes"]["tnVzBrCPName"])) 
                                rows[7] = write_cell(sheet, rows[7], 5, e["fvRsCons"]["attributes"]["tnVzBrCPName"])   
                        row = max(rows)+1
        for l3out in tn["fvTenant"]["children"]:                            #L3outs por Tenant 
            if "l3extOut" in l3out:
                row += 2
                sheet.write(row, 0, "L3Out")
                row += 1
                sheet.write(row, 0, "Name")
                sheet.write(row, 1, "VRF")
                sheet.write(row, 2, "External EPG")
                sheet.write(row, 3, "ExtEPG Subnet")
                sheet.write(row, 4, "Provide")
                sheet.write(row, 5, "Consume")
                sheet.write(row, 6, "Node Profile")
                sheet.write(row, 7, "Node")
                sheet.write(row, 8, "Router-ID")
                sheet.write(row, 9, "Int Profile")
                sheet.write(row, 10, "Interfaces")
                sheet.write(row, 11, "IP")
                sheet.write(row, 12, "Secondary IP")
                sheet.write(row, 13, "Encap")
                sheet.write(row, 14, "Dominio")
                sheet.write(row, 15, "Protocol")
                row += 1
                sheet.write(row, 0, l3out["l3extOut"]["attributes"]["name"])
                rows = [row] * 10
                for l3 in l3out["l3extOut"]["children"]:
                    if "ospfExtP" in l3:
                        write_cell(sheet, row, 15, "OSPF")          #TODO incorporar los otros portocolos de ruteo
                    if "l3extRsL3DomAtt" in l3:
                        rows[0] = write_cell(sheet, rows[0], 14, l3["l3extRsL3DomAtt"]["attributes"]["tDn"][4:])
                    if "l3extRsEctx" in l3:
                        row = write_cell(sheet, row, 1, l3["l3extRsEctx"]["attributes"]["tnFvCtxName"]) #####
                    if "l3extInstP" in l3:
                        rows[1] = write_cell(sheet, rows[1], 2, l3["l3extInstP"]["attributes"]["name"])
                        for extepg in l3["l3extInstP"]["children"]:
                            if "l3extSubnet" in extepg:
                                rows[2] = write_cell(sheet, rows[2], 3, extepg["l3extSubnet"]["attributes"]["ip"]) #TODO: Agregar Scopes
                            if "fvRsCons" in extepg:
                                rows[3] = write_cell(sheet, rows[3], 5, extepg["fvRsCons"]["attributes"]["tnVzBrCPName"])
                            if "fvRsProv" in extepg:
                                rows[4] = write_cell(sheet, rows[4], 4, extepg["fvRsProv"]["attributes"]["tnVzBrCPName"])
                    if "l3extLNodeP" in l3:
                        rows[5] = write_cell(sheet, rows[5], 6, l3["l3extLNodeP"]["attributes"]["name"])
                        for nodel3 in l3["l3extLNodeP"]["children"]:
                            if "l3extRsNodeL3OutAtt" in nodel3:
                                rows[6] = write_cell(sheet, rows[6], 7, nodel3["l3extRsNodeL3OutAtt"]["attributes"]["tDn"][9:])
                                write_cell(sheet, rows[6], 8, nodel3["l3extRsNodeL3OutAtt"]["attributes"]["rtrId"])
                            if "l3extLIfP" in nodel3:
                                rows[7] = write_cell(sheet, rows[7], 9, nodel3["l3extLIfP"]["attributes"]["name"])
                                for l3int in nodel3["l3extLIfP"]["children"]:
                                    if "l3extRsPathL3OutAtt" in l3int:      #REVISAR esta parte por diferencias en tipos de interfaces ruteadas
                                        rows[8] = write_cell(sheet, rows[8], 10, l3int["l3extRsPathL3OutAtt"]["attributes"]["tDn"][l3int["l3extRsPathL3OutAtt"]["attributes"]["tDn"].find("[")+1:-1])
                                        write_cell(sheet, rows[8], 13, l3int["l3extRsPathL3OutAtt"]["attributes"]["encap"])
                                        for l3ip in l3int["l3extRsPathL3OutAtt"]["children"]:
                                            if "l3extMember" in l3ip:
                                                rows[9] = write_cell(sheet, rows[9], 11, l3ip["l3extMember"]["attributes"]["addr"])
                row = max(rows)+1

#Access Policies
#Domain, AAEP y VLAN Pool
sheet = wb.add_sheet("Access_Policies")
row = 0
sheet.write(row, 0, "Domain")
sheet.write(row, 1, "Dom Type")
sheet.write(row, 2, "AAEP")
sheet.write(row, 3, "VLAN Pool")
sheet.write(row, 4, "Pool Type")
sheet.write(row, 5, "Range")
sheet.write(row, 6, "Description")
row += 1    
for dom in data["polUni"]["children"]: 
    if "physDomP" in dom:
        row = write_cell(sheet, row, 0, dom["physDomP"]["attributes"]["name"])
        write_cell(sheet, row, 1, "phys")
        rows = [row] * 3   
        try:                                            
            for d in dom["physDomP"]["children"]:       
                for ac in data["polUni"]["children"]:
                    if "infraInfra" in ac:
                        for vlan in ac["infraInfra"]["children"]:       #Mapeo de vlan a Dominio fisico
                            if "fvnsVlanInstP" in vlan:
                                if vlan["fvnsVlanInstP"]["attributes"]["name"] == d["infraRsVlanNs"]["attributes"]["tDn"][18:-8]:
                                    rows[0] = write_cell(sheet, rows[0], 3, vlan["fvnsVlanInstP"]["attributes"]["name"]) 
                                    for v in vlan["fvnsVlanInstP"]["children"]:
                                        if "fvnsEncapBlk" in v:
                                            rows[1] = write_cell(sheet, rows[1], 4, v["fvnsEncapBlk"]["attributes"]["allocMode"])
                                            if v["fvnsEncapBlk"]["attributes"]["from"] == v["fvnsEncapBlk"]["attributes"]["to"]:
                                                write_cell(sheet, rows[1], 5, v["fvnsEncapBlk"]["attributes"]["from"][5:])
                                            else:
                                                write_cell(sheet, rows[1], 5, v["fvnsEncapBlk"]["attributes"]["from"][5:]+v["fvnsEncapBlk"]["attributes"]["to"][4:])
                            if "infraAttEntityP" in vlan:
                                for aep in vlan["infraAttEntityP"]["children"]:
                                    if dom["physDomP"]["attributes"]["name"] == aep["infraRsDomP"]["attributes"]["tDn"][9:]:
                                       rows[2] = write_cell(sheet, rows[2], 2, vlan["infraAttEntityP"]["attributes"]["name"]) 
        except Exception as e:
            if "physDomP" in str(e):
                rows[0] = write_cell(sheet, rows[0], 3, "NA")
            if "children" in str(e): 
                rows[1] = write_cell(sheet, rows[1], 4, "NA")
        row = max(rows)+1



#Save to file
wbname = path.split("/")[-1] + ".xls"
print("The file was generated:  {}".format(wbname))
wb.save(path + ".xls")
