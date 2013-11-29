# -*- coding: utf-8 -*-

import re
import os
import sys
import win32ui
import win32com.client
import win32com.client.dynamic
import string
import time
import datetime




global HTML
HTML = ['<HTML>','<HEAD>','</HEAD>','<BODY>','</BODY>','</HTML>']

class TABLE(object):
    
    global HTML
    
    def __init__(self, table_name):
        self.table = ['</TABLE>']
        self.table.insert(0, '<TABLE NAME = "%s">' % table_name)
        self.no_of_Column = 0
        
    def done(self):
        global HTML
        BODY_index = HTML.index('</BODY>')
        HTML = HTML[:BODY_index] + self.table + HTML[BODY_index:]  
        
    def AddColumn(self, No_of_Row):
        self.no_of_Column = self.no_of_Column + 1
        TABLE_index = self.table.index('</TABLE>')
        self.table.insert(TABLE_index,'</TR>')
        self.table.insert(TABLE_index,'<TR NAME = TR%s>' % str(self.no_of_Column))
        
        TR_index = self.table.index('<TR NAME = TR%s>' % str(self.no_of_Column)) + 1
        
        No_of_Row = range(No_of_Row)
        No_of_Row.reverse()
        for i in No_of_Row:    
            self.table.insert(TR_index,'</TD>')
            self.table.insert(TR_index,'<TD NAME = TD%s_%s>' % (self.no_of_Column ,str(i + 1)))
            
            
        print "You should insert text with 'AddContent' method."
    
    def AddContent(self, TD_name, Content):
        TD_index = self.table.index('<TD NAME = %s>' % TD_name) + 1
        self.table.insert(TD_index, Content)


def Init_gConnection():
        global gConnection
        gConnection = win32com.client.Dispatch(r"ADODB.Connection")
        #strMDB = "Provider=SQLOLEDB.1;Password=moldex3d!;Persist Security Info=True;User ID=sa;Initial Catalog=Fogbugz;Data Source=192.168.3.155" #測試Server
        strMDB = "Provider=SQLOLEDB.1;Password=admin;Persist Security Info=True;User ID=sa;Initial Catalog=Fogbugz1;Data Source=192.168.130.41"
        gConnection.Open(strMDB)

        return gConnection        


def Connect_To_DB(SQL_Command):
        #Input : SQL command
        #Output : recordset
        #將cnnection new 出來連接到資料庫
        Conn = Init_gConnection()
        cm = win32com.client.Dispatch(r"ADODB.Command")
        cm.ActiveConnection = Conn
        cm.CommandType = 1#adCmdText     #http://msdn2.microsoft.com/en-us/library/ms962122.aspx
        cm.ActiveConnection.CursorLocation = 3 #static 可以使用 RecortCount 屬性
        cm.CommandText = SQL_Command
        cm.Parameters.Refresh()
        cm.Prepared = True
        (rs1, result) = cm.Execute()
        #return rs1.Fields.Item(0).Value
        #Output_Excel(rs1, rs1.RecordCount)
        return rs1, rs1.RecordCount
    
    




with open('PIC.html', 'w') as page:
    HEAD_index = HTML.index('</HEAD>')
    BODY_index = HTML.index('</BODY>')
    
    
    T = TABLE("T")
    
    SQL = """
    Select Plugin_14_CustomBugData.moduleg15, Bug.ixBug, Bug.sTitle -- Bug.ixBug, Bug.ix
    from TagAssociation inner join Bug on TagAssociation.ixBug = Bug.ixBug 
                        inner join Plugin_14_CustomBugData on TagAssociation.ixBug = Plugin_14_CustomBugData.ixBug
    where TagAssociation.ixTag = 287 and Bug.ixStatus = 1 
    order by Plugin_14_CustomBugData.moduleg15
    
    """
    Result, ResultCount = Connect_To_DB(SQL)
    
    i = 0
    temp = ""

    while not Result .eof:
        i = i + 1
        """
        Result.Fields.Item("moduleg15").Value
        Result.Fields.Item("ixBug").Value
        Result.Fields.Item("sTitle").Value
        """
        T.AddColumn(3)
        if temp == Result.Fields.Item("moduleg15").Value:
            pass
    
        else:
            i = i + 1
            T.AddColumn(3)
            T.AddContent("TD%s_1" % str(i) , Result.Fields.Item("moduleg15").Value)
            temp = Result.Fields.Item("moduleg15").Value

        T.AddContent("TD%s_2" % str(i) , str(Result.Fields.Item("ixBug").Value))
        T.AddContent("TD%s_3" % str(i) , Result.Fields.Item("sTitle").Value)
        
        Result.MoveNext()    
            
            
            
    """
    a = TABLE("T1")

    a.AddColumn(3)
    a.AddContent("TD1_3", "This is TD1-3")

    a.AddColumn(2)
    a.AddContent("TD2_2", "This is TD2-2")
    a.AddColumn(1)
    a.AddContent("TD3_1", "This is TD3-1")
    """
    T.done()
    print HTML
    page.writelines(HTML)
    
    
    
    
    
    