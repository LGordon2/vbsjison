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

%options flex

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
{S}{E}{L}{E}{C}{T}		return 'SELECT';
{C}{A}{S}{E}			return 'CASE';
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
"."				return '.';
<<EOF>>				return 'EOF';
.				return 'INVALID';

/lex

%start script

%%

primary_expression
	: IDENTIFIER
	| NUMBER
	| STR_CONST
	| '(' expression ')'
	;

script
	: statement_list EOF
	;

else_if_clause
	: else_clause
	| ELSEIF expression THEN statement_list else_if_clause
	;

else_clause
	: ELSE statement_list
	;

case_list
	: CASE expression statement_list
	| case_list CASE expression statement_list
	| case_list CASE ELSE statement_list
	;

selection_statement
	: IF expression THEN statement_list END IF
	| IF expression THEN statement_list else_if_clause END IF
	| SELECT CASE expression case_list END SELECT
	;

for_statement
	: FOR IDENTIFIER '=' NUMBER TO factor expr_list NEXT
	;

statement_list
	: statement
	| statement_list statement
	;

declaration_statement
	: DIM IDENTIFIER
	;

expression
	: assignment_expression
	;

assignment_expression
	: logical_or_expression
	| unary_expression '=' assignment_expression
	;

logical_or_expression
	: logical_and_expression
	| logical_or_expression OR logical_and_expression
	;

logical_and_expression
	: equality_expression
	| logical_and_expression AND equality_expression
	;

equality_expression
	: relational_expression
	| equality_expression '=' relational_expression
	| equality_expression NE_OP relational_expression
	;

relational_expression
	: additive_expression
	| relational_expression '<' additive_expression
	| relational_expression '>' additive_expression
	| relational_expression LE_OP additive_expression
	| relational_expression GE_OP additive_expression
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

unary_expression
	: postfix_expression
	;

postfix_expression
	: primary_expression
	;

statement
	: declaration_statement
	| expression
	| selection_statement
	;
