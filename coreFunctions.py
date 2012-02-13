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



import uno

def initializeLibreOfficeConnection():
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
    return desktop

def createNewDocument(desktop):
    document = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())
    return document

def saveDocument(document,filepath):
    document.storeAsURL(filepath,())
    
def openDocument(desktop,filepath):
    document = desktop.loadComponentFromURL(filepath,"_blank", 0, ())
    return document

def getTextCursor(document):
    cursor = document.Text.createTextCursor()
    return cursor

def insertText(text,document,cursor):
    document.Text.insertString(cursor, text, 0)
    
def insertParaSpacing(document,cursor):
    document.Text.insertString(cursor, "\n\n\t", 0)
    
def changeFont(fontName,cursor):
    cursor.setPropertyValue("CharFontName", fontName)

def changeFontSize(size,cursor):
    cursor.setPropertyValue("CharHeight", size)
    
    
    
    
    