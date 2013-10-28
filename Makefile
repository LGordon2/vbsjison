test: vbs.js
	node vbs.js test.vb

vbs.js: vbs.jison
	jison vbs.jison
