
pg-desc: ./data/Course\ Booklet\ 2019/*_PG_CourseDescription.xlsx
	./upd-one-by-one pg
pg: main.py parts/*.tex
	sed -i 's/Crimson/Blue/g' parts/pre-doc.tex
	sed -i 's/front-page-ug.pdf/front-page-pg.pdf/g' ./parts/pre-doc.tex
	sed -i 's/BTech \\\& BDes/MTech, MSc \\\& PhD/g' ./parts/pre-doc.tex
	python3 ./main.py print-all PG > pg.tex
	lualatex pg.tex
	lualatex pg.tex

ug-desc: ./data/Course\ Booklet\ 2019/*_UG_CourseDescription.xlsx
	./upd-one-by-one ug
ug: main.py parts/*.tex
	sed -i 's/Blue/Crimson/g' parts/pre-doc.tex
	sed -i 's/front-page-pg.pdf/front-page-ug.pdf/g' ./parts/pre-doc.tex
	sed -i 's/MTech, MSc \\\& PhD/BTech \\\& BDes/g' ./parts/pre-doc.tex
	python3 ./main.py print-all UG > ug.tex
	lualatex ug.tex
	lualatex ug.tex

test:main.py styles/*.sty 
	python3 main.py > test.tex
	lualatex test.tex

clean:
	rm *.log *.out *.toc *.pdf *.aux
