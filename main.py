#import xlrd
import json
import csv
from decimal import Decimal

from os import walk
course_details_filelist = ['./data/UG/'+y for y in [x[2] for x in walk('./data/UG')][0]]
UG_curr = ['./data/UG/'+y for y in [x[2] for x in walk('./data/UG')][0]]
M1yr_curr = ['./data/PG/MTech-1yr/'+y for y in [x[2] for x in walk('./data/PG/MTech-1yr')][0]]
M2yr_curr = ['./data/PG/MTech-2yr/'+y for y in [x[2] for x in walk('./data/PG/MTech-2yr')][0]]
M3yr_curr = ['./data/PG/MTech-3yr/'+y for y in [x[2] for x in walk('./data/PG/MTech-3yr')][0]]
#print(course_details_filelist)

#book = xlrd.open_workbook('test-dat1.xls')
#sh1 = book.sheet_by_index(0)

#import pytablewriter as ptw
#from pytablewriter.style import Style
from openpyxl import load_workbook
import sqlite3


no_cap_list = ['CS', 'CY', 'II', 'in', 'and', '2d', 'for', 'a', 'HT', 'MT', 
        'is', 'of', 'to', 'the', 'PH', 'DSP', 'EE', 'LA', 'CA', 'FEM', 'CFD', 'IC',
        'LTE-4G', 'MAC', 'AI', 'ML', 'GIS', '-I', '-II', 'MA', 'BM', 'BO', 'LA/CA', 
        'CMOS', 'AC', 'DC', 'MOS', 'VLSI', 'AdS', 'CFT', 'MHD',  ]
def capitals(s):
    res = ''
    for w in s.split(' '):
        if w not in no_cap_list: w = w.capitalize()
        res += w + ' '
    #print(res)
    return res

import re
tex_conv = {
        '&': r'\&',
        '%': r'\%',
        '$': r'\$',
        '#': r'\#',
        '_': r'\_',
        '{': r'\{',
        '}': r'\}',
        '~': r'\textasciitilde{}',
        '^': r'\^{}',
        '\\?P<symb>[a-zA-Z]': '\\{\\g<symb>}',
        '<': r'\textless{}',
        '>': r'\textgreater{}',
        }
tex_regex = re.compile('|'.join(re.escape(key) for key in sorted(tex_conv.keys(), key = lambda item: - len(item))))
def tex_escape(text):
    """
        :param text: a plain text message
        :return: the message escaped to appear correctly in LaTeX
    """
    return tex_regex.sub(lambda match: tex_conv[match.group()], text)

def sanitize(code, name, credits, segments, pre_req, syl):
    if code is not None: code = code.replace(' ', '').upper().strip()
    if name is not None: name = capitals(name).replace('&', 'and').strip()
    try: credits = float(credits)
    except: credits = 1.0
    segments = str(segments).strip()
    if segments == None or segments == 'None': segments = ''
    if pre_req == None or pre_req == 'None': pre_req = '' # pre_reqs are not processed as codes, should be.
    else: pre_req = pre_req.replace('&', 'and')
    if syl is not None: syl = syl.replace('&', 'and').replace('\\\\ [', '[') #tex_escape(tex_escape(syl))
    return code, name, credits, segments, pre_req, syl

def get_course_db(dept):
    con = sqlite3.connect('./courses.db')
    cur = con.cursor()
    #check if it is in the db already
    if dept=='EP': dept = 'PH'
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
        credits        float NOT NULL,
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
                sem, c, n, cd, seg, pre, syl, rem, g_rem = [r[i].value for i in range(9)] 
                c, n, cd, seg, pre, syl = sanitize(c, n, cd, seg, pre, syl)
                cur.execute(ins, [sem, c, n, cd, seg, pre, syl, rem, g_rem])
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
    print("""\\newpage""")
    print("""\\textbf{%s} Course Description""" % dept)
    table_pre = "\\columnratio{0.25} \\setlength{\\columnsep}{1em} \\begin{paracol}{2}"
    print(table_pre)
    row_style = "\\describe{%s}{%s}{%s}{%s} \\switchcolumn \\vspace{0mm} \\syllabus{%s} \\switchcolumn[0]*"
    q = """SELECT code, name, credits, segments, pre_req, syllabus FROM %s_courses""" % dept
    rows = cur.execute(q).fetchall()
    for row in rows: 
        code, name, credits, segments, pre_req, syl = row
        code, name, credits, segments, pre_req, syl = sanitize(code, name, credits, segments, pre_req, syl)
        credits = Decimal(credits)
        if code[2:] == 'XXXX': continue
        if pre_req != None and pre_req != '': pre_req = "\\triangleright "+pre_req
        print(row_style % (code, credits, name, pre_req, syl))
    table_post = "\\end{paracol}"
    print(table_post)
    #print("".join(post))

seg_dict = {}
def get_segment_line(seg):
    width = 6
    try: return seg_dict[seg]
    except: 
        template = "{\\hspace{%d mm} \\begin{tcolorbox}[colback=lightblue,boxsep=0.3mm,top=0pt,bottom=0pt,width=%d mm,boxrule=0.1mm]{\\begin{center}{\\textbf{\\tiny{%s}}}\\end{center}}\\end{tcolorbox} \\hspace{%d mm}}"
        if seg != None and seg != '':
            try: start, end = [float(x) for x in seg]
            except: return seg_dict['blank']
            frontspace = (start - 1)*width
            endspace = (width - end)*width
            midspace = (end-start+1)*width
            tbox = template % (frontspace, midspace, '%s', endspace)
            seg_dict[seg] = tbox
            return tbox
        else:
            try: return seg_dict['blank']
            except:
                seg_dict['blank'] = "{\\tiny{%s}}"
                return seg_dict['blank']

def gen_curriculum(dept, sheet, title, display_seg=True):
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
    
    #print("\\hspace{-1cm}")
    print("""\\textbf{%s %s}"""%(deptname, title))
    #table_pre = """\\begin{longtable}{@{\\extracolsep{\\fill}}rllr@{}} \
    table_pre = """\\setlength\\LTleft{0pt} \\setlength\\LTright{0pt} \n \\begin{longtable}{@{\\extracolsep{\\fill}}rllr@{}} \
            \\textcolor{Grey}{\\textbf{Cred.}} & \\textcolor{Grey}{\\textbf{Code}} & \
    \\textcolor{Grey}{\\textbf{Course Title}} & \\textcolor{Grey}{\\textbf{Segments}} \\\\ \
    \\toprule \\endhead"""
    table_pre1 = """\\setlength\\LTleft{0pt} \\setlength\\LTright{0pt} \n \\begin{longtable}{@{\\extracolsep{\\fill}}rll@{}} \
            \\textcolor{Grey}{\\textbf{Cred.}} & \\textcolor{Grey}{\\textbf{Code}} & \
    \\textcolor{Grey}{\\textbf{Course Title}} \\\\ \
    \\toprule \\endhead"""
    if display_seg: print(table_pre)
    else: print(table_pre1)
    row_style = "\\currformat{%s}{%s}{%s}"
    #row_style = "\\currformat{%s}{%s}{%s}{\\{black}{lightgreen}{\\tiny{\\hspace{1cm}%s\\hspace{1cm}}}}"
    #row_style = "\\currformat{%s}{%s}{%s}{\\colorbox{lightgreen}{\\framebox(100,2){\\tiny{\\hspace{1cm}%s\\hspace{1cm}}}}}"
    row_style1 = "\\surrformat{%s}{%s}{%s}"

    q = """SELECT name, credits, segments FROM %s_courses WHERE code='%s'"""
    g_remarks = ''
    sem, t1, t2 = None, Decimal(0), Decimal(0) # t1 = total credits, t2 = credits for a semester.
    for r in sheet.iter_rows(min_row=2):
        if sem != r[0].value:
            if sem != None:
                if display_seg: print("\\textbf{%s} & Total & \\\\"%t2)
                else: print("\\textbf{%s} & Total \\\\"%t2)
            t1 += t2
            sem, t2 = r[0].value, Decimal(0)
            if sem != None: 
                if display_seg: print(" \\multicolumn{2}{l}{\\textbf{Semester %s}} & & \\\\" % sem)
                else: print(" \\multicolumn{2}{l}{\\textbf{Semester %s}} & \\\\" % sem)
        code = r[1].value
        dept = code[:2].upper()
        #print(code, dept)
        if code[2:] == 'XXXX': name, credits, segments = r[2].value, r[3].value, r[4].value
        else: 
            try: name, credits, segments = cur.execute(q % (dept, code)).fetchall()[0]
            except: name, credits, segments = r[2].value, r[3].value, r[4].value
        #print(code, name, credits, segments)
        code, name, credits, segments, pre_req, syl = sanitize(code, name, credits, segments, '', '')
        credits = Decimal(credits)
        #segments = str(segments).strip()
        #if segments == None or segments == 'None': segments = ''
        t1 += credits
        t2 += credits
        #print(row_style % (code, name.title(), credits, segments))
        if display_seg: 
            tbox = get_segment_line(segments)
            print((row_style+tbox) % (credits, code, name, segments))
        else: print(row_style1 % (credits, code, name))
        if r[6].value != None: g_remarks += r[6].value
    if t2 != 0: print("\\textbf{%s} & Total & \\\\"%t2)
    #print("\\textbf{%d} & Grand Total & \\\\"%t1)
    table_post = "\\hline \\end{longtable}"
    print(table_post)
    print(g_remarks+" \n")
    print("""\\vspace{6mm}""")
    #print("".join(post))

def print_part(f): print(''.join(open(f).readlines()))

import sys
if __name__ == "__main__":
    def get_deptname(f): return f.split('-')[0].split('/')[-1].upper()
    if len(sys.argv) > 1:
        if sys.argv[1] == "update":
            for x in course_details_filelist:
                #print(x)
                deptname = get_deptname(x) 
                wb = load_workbook(x)
                desc = wb.get_sheet_by_name("course-descriptions")
                update_dept_cdesc(deptname, desc)
        elif sys.argv[1] == "print-doc":
            print_part('./parts/pre-doc.tex')
            print("\\chapter*{2 B.Tech Course Curriclum}\\stepcounter{chapter}\\addcontentsline{toc}{chapter}{2 B.Tech Course Curriculum}")
            import proc_list
            for s in ['curriculum', 'baskets', 'minor', 'honors', 'double-major']:
                for f in proc_list.ug_plist:
                    deptname = get_deptname(f[0]) 
                    wb = load_workbook(f[0])
                    if s in f[1]:
                        sheet = wb.get_sheet_by_name(s)
                        gen_curriculum(deptname, sheet, "UG %s Curriculum" % (capitals(s)), display_seg=s[:4]=="curr")
            print("\\chapter*{3 Course Descriptions}\stepcounter{chapter}\\addcontentsline{toc}{chapter}{3 Course Descriptions}")
            for f in proc_list.ug_plist:
                deptname = get_deptname(f[0]) 
                gen_course_description(deptname)
            print_part('./parts/post-doc.tex')
        elif sys.argv[1] == "print-one":
            # prints for one department that is at front in ug_plist
            print_part('./parts/pre-doc.tex')
            import proc_list
            for i in [1]:
                f = proc_list.ug_plist[0]
                deptname = get_deptname(f[0]) 
                wb = load_workbook(f[0])
                for s in f[1]:
                    sheet = wb.get_sheet_by_name(s)
                    disp_seg = False
                    if s[:4] == 'curr': disp_seg = True
                    gen_curriculum(deptname, sheet, s.capitalize(), display_seg=disp_seg)
                gen_course_description(deptname)
            print_part('./parts/post-doc.tex')
            
    
    


