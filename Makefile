
test:main.py styles/*.sty 
	python3 main.py > test.tex
	lualatex test.tex

