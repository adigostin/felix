
; This file is auto-generated when at least one file in project
; has BuildTool set to Assembler (sjasmplus). Changes made manually
; to this file be lost next time it is generated.
;
; It contains project settings that need to be passed to the assembler,
; which the assembler only supports as directives in source code, not
; as command-line parameters.
;
; In the list of files passed to the assembler, this file comes first.

	; This comes from the platform selected in the drop-down in the Standard toolbar.
	; (Currently only ZX Spectrum 48K is supported.)
	DEVICE %DEVICE%

	; This comes from the Project Options -> Debugging -> LoadAddress.
	ORG %LOAD_ADDR%
