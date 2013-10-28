test: vbs.js
	node vbs.js sample.vb

vbs.js: vbs.jison
	jison vbs.jison
