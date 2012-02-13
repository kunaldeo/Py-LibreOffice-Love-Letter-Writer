#   Py LibreOffice Love Letter Writer 0.1
#  Copyright YEAR 2011  Kunal Deo  (email : kunaldeo@gmail.com)

#   This program is free software; you can redistribute it and/or modify
#   it under the terms of the GNU General Public License, version 2, as 
#   published by the Free Software Foundation.

#   This program is distributed in the hope that it will be useful,
#   but WITHOUT ANY WARRANTY; without even the implied warranty of
#   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#   GNU General Public License for more details.

#   You should have received a copy of the GNU General Public License
#   along with this program; if not, write to the Free Software
#   Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

import coreFunctions
from os.path import expanduser 
from PyQt4 import QtGui 

def write(self,senderName,receiverName,senderGender,receiverGender,relationship,fileName):
    desktop = coreFunctions.initializeLibreOfficeConnection()
    document = coreFunctions.createNewDocument(desktop)
    cursor = coreFunctions.getTextCursor(document)
    coreFunctions.changeFont('Courier',cursor)
    coreFunctions.changeFontSize(15,cursor)
    coreFunctions.insertText("My Dear "+relationship+" "+receiverName,document,cursor)
    coreFunctions.insertParaSpacing(document,cursor)
    coreFunctions.insertText("I am looking forward to holding you once again in my loving arms. You are beautiful. You are mine.",document,cursor)
    coreFunctions.insertText("\n\n",document,cursor)
    coreFunctions.insertText("From a loving and longing heart",document,cursor)
    coreFunctions.insertText("\n\n",document,cursor)
    coreFunctions.insertText(senderName,document,cursor)
    
    #Save the file
    file = 'file://'+expanduser('~')+'/'+ fileName
    coreFunctions.saveDocument(document,file)
    
    QtGui.QMessageBox.warning(self,"Information","Letter has been saved at "+file,QtGui.QMessageBox.Ok)

    
    
    