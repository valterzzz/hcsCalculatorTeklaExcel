# hcsCalculatorTeklaExcel

A Tekla Open Api application with a custom GUI, that can be used to automatically calculate Hallow Core Slabs drawn in Tekla Structures. Uses premade Excel for calculations and SQLite database for storing information.

Disclaimer - This is a prototype and should be used only as basis for full production level application.

How to use -

Insert the folder hcsCheckerData (with the excel inside it) inside your Tekla Structures Model folder.
Insert the profile files in profile catalog using Tekla Structures.
Insert the rebar file in rebar catalog using Tekla Structures. (Might need a TS Restart)
Insert test strand patterns from testStrandPatterns folder in model folder\attributes folder.
Open the hcsCalculatorTeklaExcel.sln file using Visual Studio if you want to be able to eddit the code and run it with Visual Studio.
Open the /hcsCalculatorTeklaExcel/bin/Debug/hcsCalculatorTeklaExcel.exe file to use the application as standalone (Works only with Tekla Structures 2020).
Youtube link - TBC..

Known limitations -

*) HCS ends are straight (no angles) 
*) HCS are matching global Z plane (no inclinations) 
*) No strands in Tekla class 999999 
*) Max One layer top reinfrocement strands (above HCS Height/2) 
*) Max Two layers bottom reinforcement strands (below HCS Height/2)

Limitations are mostly to do with the Excel spread sheet quality and since it is meant as a place holder, there is no sensable reason to fix them up.

If you have any questions reach out to me on Linkedin - https://www.linkedin.com/in/valters-s%C4%93j%C4%81ns-84a274101/