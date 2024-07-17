Macro for Solid Edge ST, the intention is to use it right after importing a model from STEP, IGES, SAT or any other format, to heal-optimize and recognize some features (only holes at this moment), both in part and sheet metal mode. It also move the model into synchronous mode, hides the coordinate system icon, hide the reference planes, make a zoom fit and ends in ordered mode.

How-To:
You must follow these instructions two times, in Part and SheetMetal Environment.

Step 1: Make a folder named "Macros" in "C:\Program Files\Siemens\" (or where you want) and place here the Macro ---> "C:\Program Files\Siemens\Macros\OptimizarPlus.exe"

Step 2: Open a new Part/SheetMetal document, right click somewhere in the ribbon and select "Customize the ribbon".

Step 3: Select Macros in the drop-down list "Choose commands from".

Step 4: Click "Browse" at the bottom/left and select the folder created in the Step 1.

Step 5: Hightlight OptimizarPlus.exe in the left side and select the place you want the icon of the Macro in the right side and click in Add. You can expand, move up, down, rename, etc until the icon is placed where you want.

Step 6: Close and then save changes. Remember you must do it two times in Part and SheetMetal Environment.

Limitations:
- It's intented to use it right after importing a model from STEP, IGES, SAT or any other format. Not for parts designed within SolidEdge, these have nothing to heal-optimize or recognize.
- You must have only one document opened and one instance of SolidEdge running.
- You can't use it when editing a part in a assembly.

Other option:
For the optimization of many files or large projects you can use the project https://github.com/rmcanany/SolidEdgeHousekeeper that does this and many more things. Check it out.

Enjoy
