**Install:**

1. Build with C++ VS 2022, "**Release** x64" config
   
2. Copy "CustomizeToolbar.dll" to
  %ProgramFiles%\Notepad++\plugins_CustomizeToolbar\ folder and rename to
  "_CustomizeToolbar.dll"

3. Launch
  Notepad++, goto Plugins~"Customize Toolbar" menu and click
  "Custom Buttons", and close Notepad++ (NPP)
  
4. Goto
  %AppData%\Roaming\Notepad++\plugins\config\

5. The files, "CustomizeToolbar.btn" (Custom Button config text file)
  & "CustomizeToolbar.dat" (your personal Customized Toolbar
  settings are saved here), are auto generated

6. For Custom Buttons,
    modify the CustomizeToolbar.btn file entering NPP menu
  paths (comma-deliminated, up to 4-levels deep), and add icon filepath(s)
  that exist on your system (up to 3 icons).
      
      a. CustomizeToolbar.btn line example with 3 icons:
   
             Plugins,Compare,Navigation Bar,,standard-3.bmp,fluentlight-3.ico,fluentdark-3.ico

   To comment-out line, use   ";"

**Notes:**

1. This  is just @dave-user 's plugin,
   I only add and tweaked a couple lines and
  added files for Visual Studio 2022 to compile.

        a. I have only a miniscule knowledge of C++ or Notepad++ plugins so I am unable to maintain this
          further. I truely hope someone is able to continue to update
          CustomizeToolbar
   
        b. Suggestions Welcome, but include instructions.
  
3. All  Icons should be saved as 24-bit Bitmap (.bmp) and renamed to .ico.
  This is generally the universal case for icons used in Notepad++
  (Notepad-Plus-Plus). This can be done with MSPaint.exe by default.
