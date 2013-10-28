/* lexical grammar */
%lex

A	[Aa];
B	[Bb];
C	[Cc];
D	[Dd];
E	[Ee];
F	[Ff];
G	[Gg];
H	[Hh];
I	[Ii];
J	[Jj];
K	[Kk];
L	[Ll];
M	[Mm];
N	[Nn];
O	[Oo];
P	[Pp];
Q	[Qq];
R	[Rr];
S	[Ss];
T	[Tt];
U	[Uu];
V	[Vv];
W	[Ww];
X	[Xx];
Y	[Yy];
Z	[Zz];

%%
\s+				/* skip whitespace */
\'.*				/* ignore comments */
REM.*				//
Dim				return 'DIM';
Option				return 'OPTION';
Explicit			return 'EXPLICIT';
{C}{A}{L}{L}			return 'CALL';
{F}{O}{R}			return 'FOR';
{N}{E}{X}{T}			return 'NEXT';
{T}{O}				return 'TO';
{F}{U}{N}{C}{T}{I}{O}{N}	return 'FUNCTION';
{E}{L}{S}{E}{I}{F}		return 'ELSEIF';
{I}{F}				return 'IF';
{T}{H}{E}{N}			return 'THEN';
{E}{L}{S}{E}			return 'ELSE';
{E}{N}{D}			return 'END';
{O}{R}				return 'OR';
{A}{N}{D}			return 'AND';
[A-Za-z][A-Za-z0-9_]*		return 'IDENTIFIER';
[0-9]+				return 'NUMBER';
\"(\\.|[^\\"])*\"		return 'STR_CONST';
"<="				return 'LE_OP';
">="				return 'GE_OP';
"<>"				return 'NE_OP';
"="				return '=';
">"				return '>';
"<"				return '<';
"("				return '(';
")"				return ')';
","				return ',';
"*"				return '*';
"/"				return '/';
"+"				return '+';
"-"				return '-';
<<EOF>>				return 'EOF';
.				return 'INVALID';

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
	: IDENTIFIER '=' add_expr
	;

add_expr 
	: term | add_expr add_op term;

term
	: factor | term mult_op factor;

factor
	: IDENTIFIER
	| NUMBER
	| function_call
	;

add_op
	: '+' | '-';

mult_op 
	: '*' | '/';

arg_list
	: obj
	| arg_list ',' obj
	;

function_call
	: IDENTIFIER '(' arg_list ')'
	;

function_def
	: FUNCTION IDENTIFIER '(' arg_list ')' expr_list END FUNCTION
	;

obj
	: STR_CONST
	| add_expr
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

if_statement
	: IF condition_list THEN expr_list END IF
	| IF condition_list THEN expr_list ELSE expr_list END IF
	| IF condition_list THEN expr_list else_if_clause END IF
	| IF condition_list THEN expr_list else_if_clause ELSE expr_list END IF
	;

for_statement
	: FOR IDENTIFIER '=' NUMBER TO factor expr_list NEXT
	;

expr
	: DIM IDENTIFIER
	| OPTION EXPLICIT
	| CALL function_call
	| if_statement
	| for_statement
	| function_def
	| assignment
	;
