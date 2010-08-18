if (%1) == () goto usage

SET VBP=orders.vbp
if %1. == Torge. SET VBP=Torge.vbp

cd %1
"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE" /m %VBP%
cd ..

goto done

:usage
echo **********************************************
echo * Usage: %~nx0 project_folder
echo *     project_folder:: odll|torge|bay
echo **********************************************


:done