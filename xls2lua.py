#coding='utf-8'

import os.path
import sys
import xlrd


HDLC = '''-- This file is created by script!'''

class Xls2Lua():
    def __init__(self,input, output):
        self._input = input
        self._output = output
        '''
        self.pyBook = {
            sheetname1 = [
                {c1 = v1,c2=v2,...}
                ,...
            ],...
        }
        '''
        self.pyBook = {}

    def loadFile(self,filename):
        if not os.path.isfile(filename):
            raise NameError, 'invalid file name ' % filename

        book = xlrd.open_workbook(filename, formatting_info = True, encoding_override = 'utf-8')

        self.pyBook = {}
        for sheet in book.sheets():
            #invalid sheet
            if sheet.nrows < 3:
                continue

            pySheet = []
            #column names
            propNames = []
            for cell in sheet.row(0):
                propNames.append(str(cell.value))

            #default value
            defaultCells = sheet.row(2)
            for ridx in xrange(3, sheet.nrows):
                row  = {}
                for cidx in xrange(sheet.ncols):
                    cell = sheet.cell(ridx,cidx)

                    #replace empty cell by default
                    if cell.ctype == xlrd.XL_CELL_EMPTY or cell.ctype == xlrd.XL_CELL_BLANK:
                        cell = defaultCells[cidx]
                    value =  Xls2Lua.format(cell,book.datemode)
                    row[propNames[cidx]] = value
                pySheet.append(row)
            self.pyBook[sheet.name] = pySheet

    def toLua(self, outfile = '-'):
        file = open(outfile,'w')
        
        file.write(HDLC + '\n\n')
        # write sheet names.
        file.write('data = {')
        for name in self.pyBook.keys():
            file.write(' %s = {},' % name)
        file.write('}\n\n')

        for sheetname, sheet in self.pyBook.items():
            file.write('data.%s = {\n' % sheetname)
            n = len(sheet)
            for i in range(n):
                row = sheet[i]
                file.write('\t{')
                for colName, value in row.items():
                    try:
                        if type(value) is int:
                            strV = '%d' % value
                        elif type(value) is float:
                            strV = '%0.8g' % value
                        else:
                            v = ("%s"%(value)).encode("UTF-8")
                            strV = '[[%s]]' % (v)
                        file.write(' %s = %s,' % (colName, strV))
                    except Exception, e:
                        raise Exception("Format string error: (sheet:%s,row:%d,column:%s) : %s"%(sheetname,i,colName,str(e)))
                file.write(' },\n')
            file.write('}\n\n')
        file.close()

    @staticmethod
    def format(cell,datemode):
        value, ctype = cell.value, cell.ctype
        if ctype == xlrd.XL_CELL_NUMBER:
            if value == int(value): #1.0 --> 1
                value = int(value)
        elif ctype == xlrd.XL_CELL_DATE:
            dateTuple = xlrd.xldate_as_tuple(value,datemode)
            # time only no date component
            if dateTuple[0] == 0 and dateTuple[1] == 0 and dateTuple[2] == 0:
                value = '%02d:%02d:%02d' % dateTuple[3:]
            # date only, no	time
            elif dateTuple[3] == 0 and dateTuple[4]	== 0 and dateTuple[5] == 0:
                value =	'%04d/%02d/%02d' % dateTuple[:3]
            else: #	full date
                value =	'%04d/%02d/%02d	%02d:%02d:%02d'	% dateTuple

        return value


    def convert(self):
        self.loadFile(self._input)
        self.toLua(self._output)

def main():
    inst = Xls2Lua('test.xls', 'test.lua')
    inst.convert()


if __name__ == '__main__':
    main()


#TODO:  1.从命令行参数运行
#       2.字段顺序固定与excel一致