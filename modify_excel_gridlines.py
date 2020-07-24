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

def updateZip(zipin, zipout):
    # generate a temp file
    tmpfd, tmpname = tempfile.mkstemp(dir=os.path.dirname(zipout))
    os.close(tmpfd)

    # create a temp copy of the archive without filename            
    with zipfile.ZipFile(zipin, 'r') as zin:
        with zipfile.ZipFile(tmpname, 'w') as zout:
            zout.comment = zin.comment # preserve the comment
            for item in zin.infolist():
                if (item.filename == "xl/worksheets/sheet1.xml"):
                    print(item.filename + " found!")
                    zout.writestr(item, zin.read(item.filename))
                else:
                    zout.writestr(item, zin.read(item.filename))

    # replace with the temp archive
    if os.path.exists(zipout):
        os.remove(zipout)
    os.rename(tmpname, zipout)

updateZip(infile, outfile)


