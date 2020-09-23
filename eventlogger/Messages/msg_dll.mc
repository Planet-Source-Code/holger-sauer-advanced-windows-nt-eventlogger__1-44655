;// Header
MessageIdTypedef = DWORD

SeverityNames=(
    Success=0x0
    Informational=0x1
    Warning=0x2
    Error=0x3
    )

LanguageNames=(Deutsch=0x407:MSG00407)

MessageId=0x0
Severity=Success
SymbolicName=W_A_SUCCESS
Language=Deutsch
%1
.

MessageId=0x1
Severity=Informational
SymbolicName=W_A_INFO
Language=Deutsch
%1
.

MessageId=0x2
Severity=Warning
SymbolicName=W_A_WARNING
Language=Deutsch
%1
.

MessageId=0x3
Severity=Error
SymbolicName=W_A_ERROR
Language=Deutsch
%1
.

MessageId=0x10
Severity=Error
SymbolicName=W_ERROR_MSG_EXT
Language=Deutsch
Fehler %1 in Modul %3 (bei %4) aufgetreten: %2
.
