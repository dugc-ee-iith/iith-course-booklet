#import xlrd
import json
import csv

from os import walk
course_details_filelist = ['./data/UG/'+y for y in [x[2] for x in walk('./data/UG')][0]]
#print(course_details_filelist)

#book = xlrd.open_workbook('test-dat1.xls')
#sh1 = book.sheet_by_index(0)

#import pytablewriter as ptw
#from pytablewriter.style import Style
from openpyxl import load_workbook
import sqlite3

def get_course_db(dept):
    con = sqlite3.connect('./courses.db')
    cur = con.cursor()
    #check if it is in the db already
    chk = """SELECT name FROM sqlite_master WHERE type='table' AND name='%s_courses';""" % dept
    chk_res = cur.execute(chk).fetchall()
    if len(chk_res)==0: 
        print("Table %s_courses not found, please upload..." % dept)
        sys.exit()
    return cur

def update_dept_cdesc(dept, sheet):
    #use like this: update_dept_cdesc('ee', wb.get_sheet_by_name("course-descriptions"))
    con = sqlite3.connect('./courses.db')
    cur = con.cursor()
    #check if it is in the db already
    chk = """SELECT name FROM sqlite_master WHERE type='table' AND name='%s_courses';""" % dept
    chk_res = cur.execute(chk).fetchall()
    if len(chk_res)==0: 
        print("Table %s_courses not found, creating..." % dept)
        # create table
        # Code is unique, name is not, as same name can be offered for UG and PG.
        c = """
        CREATE TABLE %s_courses (
        code           STRING  PRIMARY KEY UNIQUE NOT NULL,
        name           STRING  NOT NULL,
        credits        INTEGER NOT NULL,
        semester       INTEGER,
        pre_req        STRING,
        syllabus       TEXT,
        segments       STRING,
        remarks        TEXT,
        global_remarks TEXT
        );""" % dept
        cur.execute(c)
        con.commit()
    
    ins = """INSERT INTO %s_courses (semester, code, name, credits, segments, pre_req, 
            syllabus, remarks, global_remarks) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""" % dept
    chk = """SELECT code FROM %s_courses WHERE code = ?;""" % dept
    for r in sheet.iter_rows(min_row=2):
        chk_res = cur.execute(chk, [r[1].value]).fetchall()
        if len(chk_res)==0:
            try:
                cur.execute(ins, [r[0].value, r[1].value.upper(), r[2].value, 
                    int(r[3].value), r[4].value, r[5].value, r[6].value,
                    r[7].value, r[8].value])
            except Exception as e: 
                print(r, r[1].value, r[2].value)
                print(e)
    con.commit()
    con.close()

def csv2latex():
    pre = open("./styles/descript_pre.sty").readlines()
    post = open("./styles/descript_post.sty").readlines()
    print("".join(pre))
    wb = load_workbook("EE.xlsx")
    desc = wb.get_sheet_by_name("course-descriptions")
    
    ## table environments are probelmatic since they don't break between
    ## a row. Here we have lots of text in a row, which can overflow a page.
    #table_pre = "\\begin{table*}\centering \\ra{1.8} \\begin{tabular}{@{}ll@{}} \\toprule"
    #table_pre = "\\begin{longtable}{@{}ll@{}} \
    #\\textcolor{Grey}{\\textbf{Course}} & \\textcolor{Grey}{\\textbf{Syllabus}} \\\\ \ 
    #\\toprule \\endhead"
    #print(table_pre)
    #row_style = "\\describe{%s}{%s}{%s}{%s} & \\syllabus{%s} \\\\"
    ## parcolumns had problem with setting width? 
    #row_pre = "\\begin {parcolumns}[colwidths={1=0.4\\textwidth,2=0.8\\textwidth}]{2}"
    #row_style = "\\colchunk{ \\describe{%s}{%s}{%s}{%s} } \\colchunk{ \\syllabus{%s} }"
    #row_post = "\\end{parcolumns}"

    table_pre = "\\columnratio{0.2} \\setlength{\\columnsep}{1em} \\begin{paracol}{2}"
    print(table_pre)
    row_style = "\\describe{%s}{%s}{%s}{%s} \\switchcolumn \\vspace{3mm} \\syllabus{%s} \\switchcolumn[0]*"
    for r in desc.iter_rows(min_row=2):
        code, credits, segments, name, syl = r[1].value, r[3].value, r[4].value, r[2].value, r[6].value
        if segments == None or segments == 'None': segments = ''
        name = name.replace('&', 'and')
        syl = syl.replace('\\', '').replace('&', 'and')
        try: print(row_style % (code, credits, segments, name, syl))
        except: pass
    
    #table_post = "\\bottomrule \\end{tabular} \\end{table*}"
    #table_post = "\\bottomrule \\end{longtable}"
    #print(table_post)
    table_post = "\\end{paracol}"
    print(table_post)
    print("".join(post))

def gen_course_description(dept):
    #pre = open("./parts/pre_descr_table.tex").readlines()
    #post = open("./parts/post_descr_table.tex").readlines()
    #print("".join(pre))
    cur = get_course_db(dept)
    table_pre = "\\columnratio{0.2} \\setlength{\\columnsep}{1em} \\begin{paracol}{2}"
    print(table_pre)
    row_style = "\\describe{%s}{%s}{%s}{%s} \\switchcolumn \\vspace{3mm} \\syllabus{%s} \\switchcolumn[0]*"
    q = """SELECT code, credits, segments, name, syllabus FROM %s_courses""" % dept
    rows = cur.execute(q).fetchall()
    for row in rows: 
        code, credits, segments, name, syl = row
        if segments == None or segments == 'None': segments = ''
        name = name.replace('&', 'and')
        #if syl in not None: syl = syl.replace('\\', '').replace('&', 'and')
        if syl is not None: 
            syl = syl.replace('&', 'and').replace('\\', '\\textbackslash ')
        print(row_style % (code, credits, segments, name, syl))
    table_post = "\\end{paracol}"
    print(table_post)
    #print("".join(post))

def gen_curriculum(dept, sheet, title):
    #pre = open("./parts/pre_curr_table.tex").readlines()
    #post = open("./parts/post_curr_table.tex").readlines()
    #print("".join(pre))
    cur = get_course_db(dept)
    #table_pre = "\\begin{longtable}{@{}llll@{}} \
    #\\textcolor{Grey}{\\textbf{Course Code}} & \\textcolor{Grey}{\\textbf{Course Name}} & \
    #\\textcolor{Grey}{\\textbf{Credits}} & \\textcolor{Grey}{\\textbf{Segments}} \\\\ \
    #\\toprule \\endhead"
    #{@{}rlll@{}} \
    #table_pre = """\\setlength\\LTleft{0pt} \\setlength\\LTright{0pt} \\begin{longtable}{@{\extracolsep{\\fill}}rlll@{}}
    print("""\\textbf{%s UG Curriculum}"""%deptname)
    table_pre = """\\begin{longtable}{@{}rlll@{}} \
            \\textcolor{Grey}{\\textbf{Cred.}} & \\textcolor{Grey}{\\textbf{Code}} & \
    \\textcolor{Grey}{\\textbf{Course Title}} & \\textcolor{Grey}{\\textbf{Segments}} \\\\ \
    \\toprule \\endhead"""
    print(table_pre)
    row_style = "\\currformat{%s}{%s}{%s}{%s}"
    q = """SELECT credits, segments, name FROM %s_courses WHERE code='%s'"""
    sem, t1, t2 = "0", 0, 0 # t1 = total credits, t2 = credits for a semester.
    for r in sheet.iter_rows(min_row=2):
        if sem != r[0].value:
            if sem != "0": print("\\textbf{%d} &  & \\\\"%t2)
            t1 += t2
            sem, t2 = r[0].value, 0
            print(" & \\textbf{Sem. %s} & & \\\\" % sem)
        code = r[1].value
        dept = code[:2].upper()
        #print(code, dept)
        try: credits, segments, name = cur.execute(q % (dept, code)).fetchall()[0]
        except: 
            #print(r[3].value, r[4].value, r[2].value)
            credits, segments, name = int(r[3].value), r[4].value, r[2].value
        t1 += int(credits)
        t2 += int(credits)
        if segments == None or segments == 'None': segments = ''
        name = name.replace('&', 'and')
        #print(row_style % (code, name.title(), credits, segments))
        print(row_style % (credits, code, name.title(), segments))
    print("\\textbf{%d} &  & \\\\"%t2)
    #print("\\textbf{%d} & Grand Total & \\\\"%t1)
    table_post = "\\end{longtable}"
    print(table_post)
    #print("".join(post))

def print_part(f): print(''.join(open(f).readlines()))

import sys
if __name__ == "__main__":
    def get_deptname(f): return f.split('-')[0].split('/')[-1].upper()
    if len(sys.argv) > 1:
        if sys.argv[1] == "update":
            for x in course_details_filelist:
                #print(x)
                deptname = get_deptname(f) 
                wb = load_workbook(x)
                desc = wb.get_sheet_by_name("course-descriptions")
                update_dept_cdesc(deptname, desc)
        elif sys.argv[1] == "print-desc":
            print_part('./parts/pre-doc.tex')
            gen_course_description(sys.argv[2])
            print_part('./parts/post-doc.tex')
        elif sys.argv[1] == "print-curr":
            wb = load_workbook(sys.argv[2])
            sheet = wb.get_sheet_by_name("curriculum")
            deptname = get_deptname(sys.argv[2]) 
            print_part('./parts/pre-doc.tex')
            gen_curriculum(deptname, sheet, "UG Curriculum")
            print_part('./parts/post-doc.tex')
        elif sys.argv[1] == "print-doc":
            print_part('./parts/pre-doc.tex')
            dlist = [
                    './data/UG/CS-BTech.xlsx',
                    './data/UG/EE-BTech.xlsx',
                    './data/UG/BT-BTech.xlsx']
                    #'./data/UG/CS-BTech.xlsx', 
                    #'./data/UG/CHE-BTech.xlsx', 
                    #'./data/UG/EE-BTech.xlsx', 
            for f in dlist:
                deptname = get_deptname(f) 
                wb = load_workbook(f)
                sheet = wb.get_sheet_by_name("curriculum")
                gen_curriculum(deptname, sheet, "UG Curriculum")
            for f in dlist:
                deptname = get_deptname(f) 
                gen_course_description(deptname)
            print_part('./parts/post-doc.tex')
    
    


