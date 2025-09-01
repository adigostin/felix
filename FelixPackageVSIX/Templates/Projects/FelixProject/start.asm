
; This is a sample application that prints some text using the ROM routines.
;
; By default, building is done using SjASMPlus to a .sna file.
; The load address and launch address are specified in the project options,
; under the Debugging page. ; You can change these settings in the project options.
;
; If you close the Simulator window, you can open it again from
; View -> Other Windows -> Simulator.

start:
	ei ; reenable interrupts as sjasmplus disabled them in the generated .sna file

	ld a, 2
	call 1601h ; CHAN-OPEN

	ld de, text
e0:	ld a, (de)
	and a
	jr z, ret_to_basic
	rst 10h
	inc de
	jr e0

ret_to_basic:
	call sample_lib_fun
	ld bc, 55
	ret

INK                     EQU 0x10
PAPER                   EQU 0x11
FLASH                   EQU 0x12
BRIGHT                  EQU 0x13
INVERSE                 EQU 0x14
OVER                    EQU 0x15
AT                      EQU 0x16
TAB                     EQU 0x17
CR                      EQU 0x0D

text:
	db AT, 12, 2, INK, 1, PAPER, 6, BRIGHT, 1, "Hello World!", CR, 0
