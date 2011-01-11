from pyExcelerator import *
import sys
import time
import commands
import getopt
import ConfigParser
from mailler import send_mail

def read_conf_file(confFile=None):
    """ [DB_CONF] """
    global db_server
    global db_database 
    global db_user
    global db_pass
    global db_query
    
    """ [MAIL_CONF] """
    global mail_server  
    global to_mail 
    global from_mail 
    global subject 
    global text 

    """ [REPORT_CONF] """
    global outputFile 
    global linesPerSheet  
    
    cp = ConfigParser.RawConfigParser()
    cp.read(confFile)
    cp.sections()
    
    """ [DB_CONF] """
    db_server = cp.get('DB_CONF', 'db_server')
    db_database = cp.get('DB_CONF', 'db_database')
    db_user = cp.get('DB_CONF', 'db_user')
    db_pass = cp.get('DB_CONF', 'db_pass')
    db_query = cp.get('DB_CONF', 'db_query')
    
    """ [MAIL_CONF] """
    mail_server = cp.get('MAIL_CONF', 'mail_server')
    to_mail = cp.get('MAIL_CONF', 'to_mail')
    from_mail = cp.get('MAIL_CONF', 'from_mail')
    subject = cp.get('MAIL_CONF', 'subject')
    text = cp.get('MAIL_CONF', 'text')

    """ [REPORT_CONF] """
    outputFile = cp.get('REPORT_CONF', 'outputFile')
    linesPerSheet = cp.getint('REPORT_CONF', 'linesPerSheet')

def usage ():
    """ Dispay Usage """
    print "Usage:" + sys.argv[0] + " [OPTIONS]"
    print "OPTIONS:"
    print "--lines|-l n: Split output into an Workbook with sheets of n lines or less each"
    print "--output|o : output file name/pattern"
    print "--query|q : query to be executed"
    print "--conf|c : ficheiro de configuracao com os dados do mail e bd do report"
    print "--help|h : print this information"
    sys.exit(2)

    
def openExcelSheet(outputFileName):
    """ Opens a reference to an Excel WorkBook object """
    workbook = Workbook()
    return workbook

def newExcelSheet(workbook, worksheetName):
    """ Opens a reference to a new Excel Worksheet object in a Workbook """
    worksheet = workbook.add_sheet(worksheetName)
    return workbook, worksheet


def writeExcelHeader(worksheet, titleCols):
    """ Write the header line into the worksheet """

    borders = Borders()
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1

    TestNoPat = Pattern()
    TestNoPat.pattern = Pattern.SOLID_PATTERN
    TestNoPat.pattern_fore_colour = 0x32

    Alg = Alignment()
    Alg.horz = Alignment.HORZ_CENTER
    Alg.vert = Alignment.VERT_CENTER

    style = XFStyle()
    style.borders = borders
    style.pattern = TestNoPat
    style.alignment = Alg

    cno = 0
    for titleCol in titleCols:
        worksheet.write(0, cno, titleCol, style)
        cno = cno + 1
        
def writeExcelRow(worksheet, lno, columns):
    """ Write a non-header row into the worksheet """

    cno = 0
    for item in columns:
        worksheet.write(lno, cno, str(columns[item]))
        cno = cno + 1
        
def closeExcelSheet(workbook, outputFileName):
    """ Saves the in-memory WorkBook object into the specified file """
    workbook.save(outputFileName)


def validateOpts(opts):
    """ Returns option values specified, or the default if none """
    
    confFile = 'conf/report.properties'
    
    for option, argval in opts:
        if (option in ("-c", "--conf")):
            confFile = argval
        if (option in ("-h", "--help")):
            usage()
    return confFile


def main():

    attach = []
    to = []
    
    """ This is how we are called """
    try:
        opts, args = getopt.getopt(sys.argv[1:], "c:h", ["conf=", "help"])
    except getopt.GetoptError:
        usage()

    confFile = validateOpts(opts)
    
    read_conf_file(confFile)


    for mail in to_mail.split(','):
        to.append(mail)

    conn1 = MySQLdb.connect(host=db_server,
                            user=db_user,
                            passwd=db_pass,
                            db=db_database)                                                                                                                                                                   
    cu_select = conn1.cursor(MySQLdb.cursors.DictCursor)

    try:
        cu_select.execute(db_query)
        
    except MySQLdb.Error, e:
        errInsertSql = "Sql ERROR!! sql is==>%s" % (db_query)
        sys.exit(errInsertSql)
        
    result_set = cu_select.fetchall()

    workbook = openExcelSheet(outputFile)

    fno = 0
    lno = 0
    titleCols = []
    sheet_n = 0

    for row in result_set:
        if lno % linesPerSheet == 0 :
            """ if the number of lines in the current sheet is the maximum allowed create a new sheet """
            sheet_n = sheet_n + 1
            workbook, worksheet = newExcelSheet(workbook, 'Sheet ' + str(sheet_n))
            """ reset line counting """
            lno = 1
        if lno == 1 :
            """ write sheet header  """
            writeExcelHeader(worksheet, row)
            """ write second line of sheet """
            writeExcelRow(worksheet, lno, row)
        else :
            """ write line in sheet """
            writeExcelRow(worksheet, lno, row)
        lno = lno + 1

    closeExcelSheet(workbook, outputFile)
    
    attach.append(outputFile)
    
    send_mail(
              send_from=from_mail,
              send_to=to,
              subject=subject,
              text=text,
              files=attach
              )

    if commands.getstatusoutput('rm '+outputFile)[0] == 0 :
        sys.exit(0)
    else:
        print "Warning: failed to remove output file!"
        sys.exit(3)

if __name__ == "__main__":
  sys.exit(main())

