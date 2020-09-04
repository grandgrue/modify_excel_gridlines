# -*- coding: utf-8 -*-
"""
Created on Fri Jul 24 15:42:27 2020

@author: sszgrm
"""


workdir = "H:\\Daten\\Project\\Desktop-2020\\SAStest\\Gridlines"
infile = workdir + "\\" + "BAU506T5061_Fertigerstellte_Gebaeude_nach-Gebaudeart.xlsx"
outfile = workdir + "\\" + "test.xlsx"

import os
import zipfile
import tempfile
import lxml.etree as et

def updateZip(zipin, zipout):
    # generate a temp file
    tmpfd, tmpname = tempfile.mkstemp(dir=os.path.dirname(zipout))
    os.close(tmpfd)

    # create a temp copy of the archive and modify one file            
    with zipfile.ZipFile(zipin, 'r') as zin:
        with zipfile.ZipFile(tmpname, 'w') as zout:
            zout.comment = zin.comment # preserve the comment
            for item in zin.infolist():
                # only the first sheet needs to be modified
                if (item.filename == "xl/worksheets/sheet1.xml"):
                    print(item.filename + " found!")
                    sheetcontents = zin.read(item.filename)
                    xml = et.ElementTree(et.fromstring(sheetcontents))
                    #print(et.tostring(xml, pretty_print=True))
                    # The attribute showGridLines must be placed within the sheetView tag
                    e = xml.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetView')
                    for i in e:
                        print("sheetView found. Adding showGridLines=0")
                        i.set('showGridLines', "0")
                    #print(et.tostring(xml, pretty_print=True))
                    # Write the XML into a new file in the xlsx-zip
                    zout.writestr(item, et.tostring(xml))
                else:
                    # Copy the contents for all other files (not sheet1.xml)
                    zout.writestr(item, zin.read(item.filename))

    # replace with the temp archive
    if os.path.exists(zipout):
        os.remove(zipout)
    os.rename(tmpname, zipout)

updateZip(infile, outfile)


