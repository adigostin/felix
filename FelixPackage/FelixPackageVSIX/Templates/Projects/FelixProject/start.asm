
	device ZXSPECTRUM48

INK                     EQU 0x10
PAPER                   EQU 0x11
FLASH                   EQU 0x12
BRIGHT                  EQU 0x13
INVERSE                 EQU 0x14
OVER                    EQU 0x15
AT                      EQU 0x16
TAB                     EQU 0x17
CR                      EQU 0x0D
	
; This is a sample application that prints some text using the ROM routines.
; The load address and launch address are specified in the project options.
;
; Building is done with sjasmplus. You can find it here: https://github.com/z00m128/sjasmplus
;
; If you close the Simulator window, you can open it again from
; View -> Other Windows -> Simulator.

	org 8000h

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

text:
	db AT, 12, 2, INK, 1, PAPER, 6, BRIGHT, 1, "Hello World!", CR, 0
