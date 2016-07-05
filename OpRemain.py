import xlrd, xlwt, os, sys, csv
import sqlite3, subprocess
from os.path import *

paths = ['/Users/MacBookPro/Desktop/마약잔량.xls','/Users/MacBookPro/Desktop/향정잔량.xls']

def GetTableFromArgs(args):
    args = [arg for arg in args if arg.endswith('.xls')]
    rec = []
    fld = []
    if args:
        for arg in args:
            wb = xlrd.open_workbook(arg)
            sht = wb.sheet_by_index(0)
            hdr, *tup = [sht.row_values(r) for r in range(sht.nrows)]
            fld = [hdr]
            rec+=tup
        return fld + rec
    else:
        return None


def WriteTableToCSV(tbl, csvFile):
    with open(csvFile, 'w', newline='') as fp:
        wtr = csv.writer(fp)
        for row in tbl:
            wtr.writerow(row)


def SetMemoryDB():
    return sqlite3.connect(":memory:")


def CreateTableToDB(db, tbl, tbl_name):
    csr = db.cursor();
    fields = repr(tuple(tbl[0]))
    csr.execute("DROP TABLE IF EXISTS "+tbl_name)
    csr.execute("CREATE TABLE {} {}".format(tbl_name,fields))
    if len(tbl) > 1:
        csr.executemany("INSERT INTO {} VALUES ({})".format(tbl_name,','.join('?'*len(tbl[0]))), tbl[1:])


def GetQueryResult(db, query):
    csr = db.cursor()
    csr.execute(query)
    db.commit()
    if query.lower().startswith('select'):
        hdr = [fld[0] for fld in csr.description]
        return [hdr] + [list(e) for e in csr]



dupledCodes = \
{
    "미졸람 주 1mg/ml 5ml" :["7MZLX","7MZL1X"],
    "염산페치딘 주사 1ml" : ["7PETX","7PET1X"],
    "주의]용량] 구연산펜타닐 주 100mcg/2ml" : ["7FTNX","7FTN2X"],
    "주의]용량] 구연산펜타닐 주 500mcg/10ml" : ["7FTN3X","7FTN10X"]
}

specialUnits = \
{
    "mg":['7UTVX'],
    "g":['7PTTX']
}

unitTbl= \
[["약품코드","함량","단위"],
["7FRFX",  12, "ml"],
["7DIAJEX", 2, "ml"],
["7MZL15",  3, "ml"],
["7MZLX",   5, "ml"],
["7MZL1X",  5, "ml"],
["7ANPX",  12, "ml"],
["7ANP20X",20, "ml"],
["7ATB2X", 0.5,"ml"],
["7MPX",    1, "ml"],
["7PETX",   1, "ml"],
["7PET1X",  1, "ml"],
["7UTVX",   1, "mg"],
["7KTM1X",  5, "ml"],
["7FTNX",   2, "ml"],
["7FTN2X",  2, "ml"],
["7FTN3X", 10, "ml"],
["7MORX",   1, "ml"],
["7MOR15X", 2, "ml"],
["7MORPX",  5, "ml"],
["7PTTX",  0.5,"g"]]

recallTable = \
[["약품코드","환자번호","원처방일자","처방번호[묶음]","반납구분"]]


szSrcTbl = 'srcTable'
szUnitTbl = 'unitTable'
szRecallTbl = 'recallTable'
update_qry = '''UPDATE {} SET 약품명='{}' WHERE 약품코드 IN ("{}");  
'''

select_qry = \
'''SELECT 불출일자, 병동, 환자번호, 환자명, 약품명, "처방량(규격단위)", 규격단위, ROUND(집계량-총량) As "잔량(규격단위)", 
((집계량-총량)*(SELECT 함량 FROM {szTunit} WHERE 약품코드={szTsrc}.약품코드 )) || (SELECT 단위 FROM {szTunit} WHERE 약품코드={szTsrc}.약품코드) As "잔량(함량단위)"
    FROM {szTsrc} WHERE 환자명 <> "" AND 약품코드 LIKE "7%" AND INSTR("처방량(규격단위)",".")
    GROUP BY "약품명", "환자번호", "처방번호[묶음]", "원처방일자" HAVING count("처방번호[묶음]")=1 order by 약품명, 불출일자
    '''
testqry = "select * from {} WHERE 약품코드 LIKE '7%' and 환자명 <> '' ORDER BY 약품코드 ".format(szSrcTbl)
create_recallTbl_qry = \
'INSERT INTO {szTrec} SELECT 약품코드, 환자번호, 원처방일자, "처방번호[묶음]", 반납구분 FROM {szTsrc} WHERE  반납구분="반납"'.format(szTrec=szRecallTbl, szTsrc=szSrcTbl)
get_recallTbl_qry = "SELECT * FROM {}".format(szRecallTbl)
delete_qry = \
'''
DELETE FROM {szTsrc} WHERE 환자번호={szTrec}.환자번호 AND 원처방일자={szTrec}.원처방일자 AND "처방번호[묶음]" = "{szTrec}.처방번호[묶음]" AND 약품코드={szTrec}.약품코드 
'''.format(szTsrc=szSrcTbl,szTrec=szRecallTbl)


DB = SetMemoryDB()
tbl = GetTableFromArgs(paths) 
CreateTableToDB(DB,tbl,szSrcTbl)
CreateTableToDB(DB,unitTbl,szUnitTbl)

for k, v in dupledCodes.items():
    qry = update_qry.format(szSrcTbl, k, '", "'.join(v))
    print(qry)
    GetQueryResult(DB,qry)

# GetQueryResult(DB,"UPDATE {} SET 약품명='{}' where 약품코드 In ('7PETX','7PETX1')".format(szSrcTbl,"abc"))
tbl = GetQueryResult(DB,select_qry.format(szTsrc=szSrcTbl, szTunit=szUnitTbl))


#print(tbl)
testfile = '/Users/MacBookPro/Desktop/python/test.csv'
WriteTableToCSV(tbl,testfile)

    



