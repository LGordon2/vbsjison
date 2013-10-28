test: vbs.js
	node vbs.js test.vb

clean:
	rm vbs.js

vbs.js: vbs.jison
	jison vbs.jison
