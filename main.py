#import xlrd
import json
import csv

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
        c = """
    CREATE TABLE %s_courses (
    code           STRING  PRIMARY KEY UNIQUE NOT NULL,
    name           STRING  UNIQUE NOT NULL,
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
                cur.execute(ins, [r[0].value, r[1].value, r[2].value, int(r[3].value), r[4].value,
                        r[5].value, r[6].value, r[7].value, r[8].value])
            except: print(r[1].value)
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
    pre = open("./styles/descript_pre.sty").readlines()
    post = open("./styles/descript_post.sty").readlines()
    print("".join(pre))

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
        syl = syl.replace('\\', '').replace('&', 'and')
        print(row_style % (code, credits, segments, name, syl))
    table_post = "\\end{paracol}"
    print(table_post)
    print("".join(post))

def gen_curriculum(dept, sheet, title):
    pre = open("./styles/curriculum_pre.sty").readlines()
    post = open("./styles/curriculum_post.sty").readlines()
    print("".join(pre))
    
    cur = get_course_db(dept)
    table_pre = "\\begin{longtable}{@{}llll@{}} \
    \\textcolor{Grey}{\\textbf{Course Code}} & \\textcolor{Grey}{\\textbf{Course Name}} & \
    \\textcolor{Grey}{\\textbf{Credits}} & \\textcolor{Grey}{\\textbf{Segments}} \\\\ \
    \\toprule \\endhead"
    print(table_pre)
    row_style = "\\currformat{%s}{%s}{%s}{%s}"
    q = """SELECT credits, segments, name FROM %s_courses WHERE code='%s'"""
    sem, t1, t2 = "0", 0, 0 # t1 = total credits, t2 = credits for a semester.
    for r in sheet.iter_rows(min_row=2):
        if sem != r[0].value:
            if sem != "0": print(" & Total & %d & \\\\"%t2)
            sem, t2 = r[0].value, 0
            print(" & Semester %s & & \\\\" % sem)
        code = r[1].value
        dept = code[:2].lower()
        #print(code, dept)
        try: credits, segments, name = cur.execute(q % (dept, code)).fetchall()[0]
        except: continue
        t1 += int(credits)
        t2 += int(credits)
        if segments == None or segments == 'None': segments = ''
        name = name.replace('&', 'and')
        print(row_style % (code, name, credits, segments))
    table_post = "\\end{longtable}"
    print(table_post)
    print("".join(post))


import sys
if __name__ == "__main__":
    if len(sys.argv) > 1:
        if sys.argv[1] == "update":
            wb = load_workbook(sys.argv[2])
            desc = wb.get_sheet_by_name("course-descriptions")
            update_dept_cdesc(sys.argv[3], sheet)
        elif sys.argv[1] == "print-desc":
            gen_course_description(sys.argv[2])
        elif sys.argv[1] == "print-curr":
            wb = load_workbook("%s.xlsx" % sys.argv[2])
            sheet = wb.get_sheet_by_name("curriculum")
            gen_curriculum(sys.argv[2], sheet, "UG Curriculum")
    
    


