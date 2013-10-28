/* lexical grammar */
%lex

%%
\s+			/* skip whitespace */
\'.*			/* ignore comments */
REM.*			//
Dim			return 'DIM';
Option			return 'OPTION';
Explicit		return 'EXPLICIT';
[Ii][Ff]		return 'IF';
[Tt][Hh][Ee][Nn]	return 'THEN';
[Ee][Ll][Ss][Ee][Ii][Ff]return 'ELSEIF';
[Ee][Ll][Ss][Ee]	return 'ELSE';
[Ee][Nn][Dd]		return 'END';
[Oo][Rr]		return 'OR';
[Aa][Nn][Dd]		return 'AND';
[A-Za-z][A-Za-z0-9_]*	return 'IDENTIFIER';
[0-9]+			return 'NUMBER';
\"(\\.|[^\\"])*\"	return 'STR_CONST';
"<="			return 'LE_OP';
">="			return 'GE_OP';
"<>"			return 'NE_OP';
"="			return '=';
">"			return '>';
"<"			return '<';
"("			return '(';
")"			return ')';
","			return ',';
"*"			return '*';
"/"			return '/';
"+"			return '+';
"-"			return '-';
<<EOF>>			return 'EOF';
.			return 'INVALID';

/lex

%start script

%%

primary_expression
	: IDENITIFIER
	| NUMBER
	| STR_CONST
	| '(' expression ')'
	;

script
	: expr_list EOF
	;

expr_list
	: expr
	| expr_list expr
	;

assignment
	: IDENTIFIER '=' IDENTIFIER
	| IDENTIFIER '=' NUMBER
	| IDENTIFIER '=' function_call
	| IDENTIFIER '=' multiplicative_expression
	;

postfix_expression
	: primary_expression
	;

unary_expression
	: postfix_expression
	;

cast_expression
	: unary_expression
	;

multiplicative_expression
	: cast_expression
	| multiplicative_expression '*' cast_expression
	| multiplicative_expression '/' cast_expression
	| multiplicative_expression '%' cast_expression
	;

additive_expression
	: multiplicative_expression
	| additive_expression '+' multiplicative_expression
	| additive_expression '-' multiplicative_expression
	;

arg_list
	: obj
	| arg_list ',' obj
	;

function_call
	: IDENTIFIER '(' arg_list ')'
	;

obj
	: IDENTIFIER
	| constant
	; 

constant
	: NUMBER
	| STR_CONST
	;

comp_op
	: LE_OP | NE_OP | GE_OP | '<' | '>' | '=';

condition_list
	: condition
	| condition_list AND condition
	| condition_list OR condition
	;

condition
	: IDENTIFIER comp_op constant
	;

else_if_clause
	: ELSEIF condition_list THEN expr_list
	| ELSEIF condition_list THEN expr_list else_if_clause
	;

else_clause
	: ELSE expr_list
	;

expr
	: DIM IDENTIFIER
	| OPTION EXPLICIT
	| IF condition_list 
	| IF condition_list THEN expr_list else_if_clause
	| IF condition_list THEN expr_list else_if_clause else_clause
	| IF condition_list THEN expr_list else_clause
	| assignment
	;
