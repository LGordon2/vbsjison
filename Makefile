test: vbs.js
	node vbs.js example2.vb

vbs.js: vbs.jison
	jison vbs.jison
