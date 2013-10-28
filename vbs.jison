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

[ \t\b\r]+				/* skip whitespace */
\'.*				/* ignore comments */
REM.*				//
\n				return 'NEWLINE';
Dim				return 'DIM';
Option				return 'OPTION';
Explicit			return 'EXPLICIT';
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
{U}{N}{T}{I}{L}			return 'UNTIL';
{D}{O}				return 'DO';
{E}{X}{I}{T}			return 'EXIT';
{S}{U}{B}			return 'SUB';
{L}{O}{O}{P}			return 'LOOP';
{W}{H}{I}{L}{E}			return 'WHILE';
{W}{E}{N}{D}			return 'WEND';
{P}{U}{B}{L}{I}{C}		return 'PUBLIC';
{P}{R}{I}{V}{A}{T}{E}		return 'PRIVATE';
True				return 'TRUE';
False				return 'FALSE';
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
"&"				return '&';
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
	| TRUE
	| FALSE
	;

script
	: statement_list EOF
	;

else_if_clause
	: else_clause
	| ELSEIF expression THEN statement_list
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

statement_list
	: statement
	| statement_list statement 
	;

dim_statement
	: DIM argument_expression_list
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
	| additive_expression '&' multiplicative_expression
	;

unary_expression
	: postfix_expression
	;

postfix_expression
	: primary_expression
	| postfix_expression '(' ')'
	| postfix_expression '(' argument_expression_list ')'
	| postfix_expression argument_expression_list
	| postfix_expression '.' IDENTIFIER
	;

argument_expression_list
	: assignment_expression
	| argument_expression_list ',' assignment_expression
	;

do_statement
	: DO WHILE expression LOOP
	| DO WHILE expression statement_list LOOP
	| DO UNTIL expression statement_list LOOP 
	| DO UNTIL expression LOOP
	| DO LOOP
	| DO statement_list LOOP WHILE expression
	| DO statement_list LOOP UNTIL expression
	;

iteration_statement
	: FOR expression TO expression statement_list NEXT
	| WHILE expression statement_list WEND
	| WHILE expression WEND
	| do_statement
	;

jump_statement
	: EXIT DO
	| EXIT FUNCTION
	| EXIT FOR
	;

expression_statement
	: expression
	;

access_modifier
	: PUBLIC
	| PRIVATE
	;

function_definition
	: FUNCTION IDENTIFIER '(' argument_expression_list ')' statement_list END FUNCTION
	| SUB IDENTIFIER '(' argument_expression_list ')' statement_list END SUB
	| access_modifier function_definition
	;

statement
	: dim_statement
	| expression_statement
	| selection_statement
	| iteration_statement
	| jump_statement
	| function_definition
	| OPTION EXPLICIT
	| NEWLINE
	;
