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

from PyQt4 import QtCore, QtGui
import sys
import os
import subprocess
import loveLetterWriter

#import UI
from mainwindow import Ui_Form 

class Main(QtGui.QMainWindow):
   def __init__(self, parent=None):
      QtGui.QWidget.__init__(self, parent)
      
      self.ui = Ui_Form()
      self.ui.setupUi(self)
      
      # Populate default values
      self.ui.portNumber.setText("2002")
      self.ui.relationship.addItems('Wife Husband Boyfriend Girlfriend'.split())
      self.ui.senderGender.addItems('Male Female'.split())
      self.ui.receiverGender.addItems('Female Male'.split())
      self.ui.letterTone.addItems('Romantic Casual Comedy Apologetic'.split())
      self.ui.fileName.setText("lovelLetter.odt")
      
      #Connect Buttons
      QtCore.QObject.connect (self.ui.startLibreOfficeButton,QtCore.SIGNAL ("clicked()"),self.startLibreOffice)
      QtCore.QObject.connect (self.ui.writeLoveLetterButton,QtCore.SIGNAL ("clicked()"),self.writeLoveLetter)
      
   def startLibreOffice(self):
      portNumber = str(self.ui.portNumber.text())
      #commands.getoutput('soffice "--accept=socket,port='+portNumber+';urp;"')
      os.system('soffice "--accept=socket,port='+portNumber+';urp;" &')
      
      
   def writeLoveLetter(self):
      #Get current values
      senderName = str(self.ui.senderName.text())
      receiverName = str(self.ui.receiverName.text())
      senderGender = str(self.ui.senderGender.currentText())
      receiverGender = str(self.ui.receiverGender.currentText())
      relationship = str(self.ui.relationship.currentText())
      fileName = str(self.ui.fileName.text())
      loveLetterWriter.write(self,senderName,receiverName,senderGender,receiverGender,relationship,fileName)
      
def main():
   app = QtGui.QApplication(sys.argv)
   window = Main()
   window.show()
   sys.exit(app.exec_())

if __name__ == "__main__":
   main()
