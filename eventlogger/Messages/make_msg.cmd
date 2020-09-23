@echo off
mc -v -U msg_dll
rc -l 409 -r -fo msg_restable.res msg_restable.rc
link -dll -MACHINE:Ix86 -noentry -out:msg_dll.dll msg_restable.res
