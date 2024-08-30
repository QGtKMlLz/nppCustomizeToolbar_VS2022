// This file is the main part of Customize Toolbar, a plugin for Notepad++
// Copyright (C) 2011-2021 DW-dev (dw-dev@gmx.com)
// Copyright (©) 2024+     QGtKMlLz    E-mail: 3m33dkojb@mozmail.com
// Last Edit - 29 Aug 2024

// Interactions with Notepad++ and Other Plugins:
//
//  - assumes that required rebar is first rebar in window
//  - assumes that required toolbar is first toolbar in rebar
//  - subclasses window and rebar
//  - derives button string from menu string
//  - workaround for Spell-Checker plugin
//  - workaround for WebEdit plugin
//  - workaround for Python Script plugin
//  - assigns temporary command identifiers to custom buttons until NPPN_READY received
//  - traps RB_SETBANDINFO message (fMask == 0x0270) to detect icons changed by Notepad++
//  - updates button states from menu states
//  - sends TB_SETMAXTEXTROWS message to force toolbar to refresh and display buttons
//  - hashes menu string and parent menu string to uniquely identify button

// CustomizeToolbar.dat File Format - for toolbar layout
//
// Versions 1.2-2.0
// number of buttons on toolbar                     XX000000
// number of buttons available                      XX000000
// first button on toolbar (cmdid or menuhash)      XXXX0000 or XXXXXXXX
// repeat for each button on toolbar                ........
// first button available (cmdid or menuhash)       XXXX0000 or XXXXXXXX
// repeat for each button available                 ........
// wrap toolbar menu item state (2.0)               00000000 or 01000000
//
// Versions 3.0-5.3
// custom buttons menu item state                   00000000 or 01000000
// wrap toolbar menu item state                     00000000 or 01000000
// number of buttons on toolbar                     XX000000
// number of buttons available                      XX000000
// first button on toolbar (cmdid or menuhash)      XXXX0000 or XXXXXXXX
// repeat for each button on toolbar                ........
// first button available (cmdid or menuhash)       XXXX0000 or XXXXXXXX
// repeat for each button available                 ........
// line number margin button state (3.10-4.2)       00000000 or 01000000
// bookmark margin button state (3.10-4.2)          00000000 or 01000000
// folder margin button state (3.10-4.2)            00000000 or 01000000

// CustomizeToolbar.btn File Format - for custom buttons
// 
// Each line is either a semi-colon followed by a comment or a custom button definition:
// menustring1,menustring2,menustring3,menustring4,standard.bmp,fluentlight.ico,fluentdark.ico
// The standard.bmp, fluentlight.ico and fluentdark.ico image file names are optional
// With Notepad++ 7.9.5 or earlier, the fluent light and fluent dark fields are ignored
// With Notepad++ 8.0 or later, the fluent dark field if omitted defaults to the fluent light field

// Example debug message:
//    _stprintf_s(g_debugBuffer, 200, TEXT("value=%i"), value);
//    MessageBox(nppData._nppHandle, g_debugBuffer, TEXT("Debug"), MB_OK);

// Include files
#include "PluginDefinition.h"
#include "menuCmdID.h"
#include "resource.h"
#include <commctrl.h>
#include <tchar.h>
#include "Shlwapi.h"
#include "versionhelpers.h"

// Constant definitions

#define IDM_EDIT_TAB2SW (IDM_EDIT + 46)
#define IDM_EDIT_SW2TAB_ALL (IDM_EDIT + 54)
#define IDM_FOCUS_ON_FOUND_RESULTS (IDM_SEARCH + 45)

#define ID_PLUGINS_CMD 22000
#define ID_PLUGINS_CMD_LIMIT_OLD 22499  /* 500 plugin buttons (with menu items) - Notepad++ 8.1.2 or earlier */
#define ID_PLUGINS_CMD_LIMIT_NEW 22999  /* 1000 plugin buttons (with menu items) - Notepad++ 8.1.3 or later */

#define ID_PLUGINS_CMD_DYNAMIC 23000
#define ID_PLUGINS_CMD_DYNAMIC_LIMIT 24999  /* 2000 plugin buttons (without menu items) */

#define ID_CMD_CUSTOM 26000
#define ID_CMD_CUSTOM_LIMIT 26099  /* 100 custom buttons */

#define MAXSIZE 300  /* maximum size of field (menu string or file name) - menu string can contain file name (260) plus a few more characters */

#define HASHFLAG 0x80000000

// Data declarations

TCHAR g_debugBuffer[200];

extern HANDLE g_hModule;
HMENU g_hMainMenu;

int g_nppVersion;
int g_id_plugins_cmd_limit;
int g_rebarBandInfoSize;

WNDPROC g_origWindowProc, g_origRebarProc;

TBBUTTON g_tbButtons[300];  /* 300 buttons in total - built-in buttons, plugin buttons, dynamic plugin buttons and custom buttons */
int g_buttonsAvailable;
int g_customButtonsState;
int g_wrapToolbarState;

TCHAR g_customMenuStrings[100][4][MAXSIZE];  /* 100 custom buttons, 4 menu strings per button */
int g_customButtonsCount;

// Function declarations

void addAdditionalButton(int bitmapName, int iconName, int idCmd);
DWORD WINAPI afterNppReadyDelayed(LPVOID lpParam);
LRESULT APIENTRY subclassRebarProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
LRESULT APIENTRY subclassWindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
DWORD WINAPI handleWindowResize(LPVOID lpParam);
DWORD WINAPI handleChangedIcons(LPVOID lpParam);
DWORD WINAPI handleButtonStates(LPVOID lpParam);
void replaceTemporaryCmdIDs();
void preserveToolbarButtons();
void updateToolbarState();
void resetToolbarLayout();
void saveToolbarLayout();
void restoreToolbarLayout(bool menuStates);
void makeToolbarWrap();
void makeToolbarOverflow();
void adjustIdealSize();
void displayOverflowMenu(NMREBARCHEVRON *lpNmRebarChevron);
HBITMAP createBitmapForCustomButton(TCHAR text[]);
HICON createIconForCustomButton(TCHAR text[]);
DWORD calcButtonStringHash(TBBUTTON tbButton);
DWORD calcPluginButtonMenuHash(TBBUTTON tbButton);
int findPluginParentMenuString(HMENU hMenu, UINT idCommand, LPTSTR lpString, int maxCount);
int findCmdIDForMenuStrings(HMENU hMenu0, LPTSTR menuString0, LPTSTR menuString1, LPTSTR menuString2, LPTSTR menuString3);
void stripMenuString(LPTSTR lpString);
int getCommCtrlMajorVersion();

//
// The plugin data that Notepad++ needs
//

FuncItem funcItem[nbFunc];

//
// The data of Notepad++ that you can use in your plugin commands
//

NppData nppData;

//
// Initialize your plugin data here
// It will be called while plugin loading
//

void pluginInit(HANDLE hModule)
{
}

//
// Here you can do the clean up, save the parameters (if any) for the next session
//

void pluginCleanUp()
{
}

//
// Initialization of your plugin commands
// You should fill your plugins commands here
//

void commandMenuInit()
{
    //--------------------------------------------//
    //-- STEP 3. CUSTOMIZE YOUR PLUGIN COMMANDS --//
    //--------------------------------------------//
    // with function :
    // setCommand(int index,                      // zero based number to indicate the order of command
    //            TCHAR *commandName,             // the command name that you want to see in plugin menu
    //            PFUNCPLUGINCMD functionPointer, // the symbol of function (function pointer) associated with this command. The body should be defined below. See Step 4.
    //            ShortcutKey *shortcut,          // optional. Define a shortcut to trigger this command
    //            bool check0nInit                // optional. Make this menu item be checked visually
    //            );
    //
    
    setCommand(0, (TCHAR*)TEXT("Customize Toolbar..."), customizeToolbar, NULL, false);
    setCommand(1, (TCHAR*)TEXT("----------"), NULL, NULL, false);
    setCommand(2, (TCHAR*)TEXT("Custom Buttons"), customButtons, NULL, false);
    setCommand(3, (TCHAR*)TEXT("Wrap Toolbar"), wrapToolbar, NULL, false);
    setCommand(4, (TCHAR*)TEXT("----------"), NULL, NULL, false);
    setCommand(5, (TCHAR*)TEXT("Help - Overview"), helpOverview, NULL, false);
    setCommand(6, (TCHAR*)TEXT("Help - Custom Buttons"), helpCustomButtons, NULL, false);
    setCommand(7, (TCHAR*)TEXT("----------"), NULL, NULL, false);
    setCommand(8, (TCHAR*)TEXT("Resource Usage"), resourceUsage, NULL, false);
}

//
// Here you can do the clean up (especially for the shortcut)
//

void commandMenuCleanUp()
{
    // Don't forget to deallocate your shortcut here
}

//
// This function help you to initialize your plugin commands
//

bool setCommand(size_t index, TCHAR *cmdName, PFUNCPLUGINCMD pFunc, ShortcutKey *sk, bool check0nInit)
{
    if (index >= nbFunc)
        return false;
    
    if (!pFunc)
        return false;
    
    lstrcpy(funcItem[index]._itemName, cmdName);
    funcItem[index]._pFunc = pFunc;
    funcItem[index]._init2Check = check0nInit;
    funcItem[index]._pShKey = sk;
    
    return true;
}

//----------------------------------------------//
//-- STEP 4. DEFINE YOUR ASSOCIATED FUNCTIONS --//
//----------------------------------------------//
//
// Toolbar control functions
//

void addMenuCommands()
{
    TCHAR configPath[MAX_PATH];
    TCHAR datFilePath[MAX_PATH];
    HANDLE datFile;
    DWORD bytesRead;
    int menuHidden;
    
    // Initialize Notepad++ version number
    
    g_nppVersion = (int) SendMessage(nppData._nppHandle, NPPM_GETNPPVERSION, 0, 0);
    
    // Initialize plugin command identifier limit
    
    if (HIWORD(g_nppVersion) >= 9 || (HIWORD(g_nppVersion) == 8 && LOWORD(g_nppVersion) >= 13)) g_id_plugins_cmd_limit = ID_PLUGINS_CMD_LIMIT_NEW;
    else g_id_plugins_cmd_limit = ID_PLUGINS_CMD_LIMIT_OLD;
    
    // Initialize main menu handle
    
    menuHidden = (int) (LRESULT) SendMessage(nppData._nppHandle, NPPM_HIDEMENU, 0, false);
    g_hMainMenu = GetMenu(nppData._nppHandle);
    SendMessage(nppData._nppHandle, NPPM_HIDEMENU, 0, menuHidden);
    
    // Initialise menu state variables for Custom Buttons and Wrap Toolbar
    
    g_customButtonsState = g_wrapToolbarState = 0;
    
    SendMessage(nppData._nppHandle, NPPM_GETPLUGINSCONFIGDIR, MAX_PATH, (LPARAM) configPath);
    
    lstrcpy(datFilePath, configPath);
    lstrcat(datFilePath, TEXT("\\CustomizeToolbar.dat"));
    
    datFile = CreateFile(datFilePath, GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    if (datFile != INVALID_HANDLE_VALUE)
    {
        ReadFile(datFile, &g_customButtonsState, sizeof(int), &bytesRead, NULL);
        ReadFile(datFile, &g_wrapToolbarState, sizeof(int), &bytesRead, NULL);
    }
    
    CloseHandle(datFile);
}

void addToolbarButtons()
{
    TCHAR configPath[MAX_PATH];
    TCHAR btnFilePath[MAX_PATH];
    TCHAR bmpFilePath[MAX_PATH], icoFilePath[MAX_PATH], icodarkFilePath[MAX_PATH];
    HANDLE btnFile, bmpFile, icoFile, icodarkFile;
    TCHAR nextChar;
    DWORD bytesRead;
    TCHAR buffer[MAXSIZE*7+10];  /* custom button definition - 4 menu strings, 3 file names, plus added commas */
    TCHAR bmpFileName[MAXSIZE], icoFileName[MAXSIZE], icodarkFileName[MAXSIZE];
    HBITMAP hToolbarBmp;
    HICON hToolbarIcon, hToolbarIconDarkMode;
    toolbarIcons buttonIcon;
    toolbarIconsWithDarkMode buttonIconDM;
    bool bmpError, icoError, icodarkError;
    int i, j;

    // Add twenty-six additional buttons onto toolbar for Notepad++ built-in commands
    
    addAdditionalButton(IDB_FILE_CLOSEALLBUT, IDI_FILE_CLOSEALLBUT, IDM_FILE_CLOSEALL_BUT_CURRENT);
    addAdditionalButton(IDB_EDIT_DELETE, IDI_EDIT_DELETE, IDM_EDIT_DELETE);
    addAdditionalButton(IDB_INDENT_DECREASE, IDI_INDENT_DECREASE, IDM_EDIT_RMV_TAB);
    addAdditionalButton(IDB_INDENT_INCREASE, IDI_INDENT_INCREASE, IDM_EDIT_INS_TAB);
    addAdditionalButton(IDB_LINE_DUPLICATE, IDI_LINE_DUPLICATE, IDM_EDIT_DUP_LINE);
    addAdditionalButton(IDB_COMMENT_SET, IDI_COMMENT_SET, IDM_EDIT_BLOCK_COMMENT_SET);
    addAdditionalButton(IDB_COMMENT_CLEAR, IDI_COMMENT_CLEAR, IDM_EDIT_BLOCK_UNCOMMENT);
    addAdditionalButton(IDB_AUTO_WORDCOMPLETE, IDI_AUTO_WORDCOMPLETE, IDM_EDIT_AUTOCOMPLETE_CURRENTFILE);
    addAdditionalButton(IDB_BLANK_TRIMTRAILING, IDI_BLANK_TRIMTRAILING, IDM_EDIT_TRIMTRAILING);
    addAdditionalButton(IDB_BLANK_TABTOSPACE, IDI_BLANK_TABTOSPACE, IDM_EDIT_TAB2SW);
    addAdditionalButton(IDB_BLANK_SPACETOTAB, IDI_BLANK_SPACETOTAB, IDM_EDIT_SW2TAB_ALL);
    addAdditionalButton(IDB_SEARCH_FINDINFILES, IDI_SEARCH_FINDINFILES, IDM_SEARCH_FINDINFILES);
    addAdditionalButton(IDB_SEARCH_FINDPREV, IDI_SEARCH_FINDPREV, IDM_SEARCH_FINDPREV);
    addAdditionalButton(IDB_SEARCH_FINDNEXT, IDI_SEARCH_FINDNEXT, IDM_SEARCH_FINDNEXT);
    addAdditionalButton(IDB_SEARCH_INCREMENTAL, IDI_SEARCH_INCREMENTAL, IDM_SEARCH_FINDINCREMENT);
    addAdditionalButton(IDB_SEARCH_RESULTS, IDI_SEARCH_RESULTS, IDM_FOCUS_ON_FOUND_RESULTS);
    addAdditionalButton(IDB_SEARCH_GOTO, IDI_SEARCH_GOTO, IDM_SEARCH_GOTOLINE);
    addAdditionalButton(IDB_BOOKMARK_PREV, IDI_BOOKMARK_PREV, IDM_SEARCH_PREV_BOOKMARK);
    addAdditionalButton(IDB_BOOKMARK_NEXT, IDI_BOOKMARK_NEXT, IDM_SEARCH_NEXT_BOOKMARK);
    addAdditionalButton(IDB_BOOKMARK_CLEAR, IDI_BOOKMARK_CLEAR, IDM_SEARCH_CLEAR_BOOKMARKS);
    addAdditionalButton(IDB_ZOOM_RESTORE, IDI_ZOOM_RESTORE, IDM_VIEW_ZOOMRESTORE);
    addAdditionalButton(IDB_MOVE_MOVETOOTHER, IDI_MOVE_MOVETOOTHER, IDM_VIEW_GOTO_ANOTHER_VIEW);
    addAdditionalButton(IDB_CLONE_CLONETOOTHER, IDI_CLONE_CLONETOOTHER, IDM_VIEW_CLONE_TO_ANOTHER_VIEW);
    addAdditionalButton(IDB_VIEW_HIDELINES, IDI_VIEW_HIDELINES, IDM_VIEW_HIDELINES);
    addAdditionalButton(IDB_VIEW_FOLDALL, IDI_VIEW_FOLDALL, IDM_VIEW_TOGGLE_FOLDALL);
    addAdditionalButton(IDB_VIEW_UNFOLDALL, IDI_VIEW_UNFOLDALL, IDM_VIEW_TOGGLE_UNFOLDALL);
    // Add customize toolbar button onto toolbar
    
    addAdditionalButton(IDB_CUSTOMIZE_TOOLBAR, IDI_CUSTOMIZE_TOOLBAR, funcItem[0]._cmdID);

    if (!g_customButtonsState) return;
    // Add custom buttons onto toolbar for Notepad++ built-in commands or plugin commands (with temporary custom command identifiers)
    
    g_customButtonsCount = 0;
    
    SendMessage(nppData._nppHandle, NPPM_GETPLUGINSCONFIGDIR, MAX_PATH, (LPARAM) configPath);
    
    lstrcpy(btnFilePath, configPath);
    lstrcat(btnFilePath, TEXT("\\CustomizeToolbar.btn"));
    
    btnFile = CreateFile(btnFilePath, GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);

    ReadFile(btnFile, &nextChar, sizeof(TCHAR), &bytesRead, NULL);
    
    //DEBUG: if (bytesRead == 0) {
    //DEBUG:     MessageBox(nppData._nppHandle, TEXT("CustomizeToolbar.btn is empty or could not read the first character"), TEXT("Read Error"), MB_OK);
    //DEBUG: }
    //DEBUG: else {
    //DEBUG:     TCHAR buffer[2] = { nextChar, 0 };
    //DEBUG:     MessageBox(nppData._nppHandle, buffer, TEXT("First Character Read"), MB_OK);
    //DEBUG: }
    //DEBUG: MessageBox(NULL, nextChar == 0xFEFF ? TEXT("BOM present") : TEXT("No BOM"), TEXT("Debug"), MB_OK);

    if (bytesRead > 0 && nextChar == 0xFEFF) ReadFile(btnFile, &nextChar, sizeof(TCHAR), &bytesRead, NULL);  /* read another character if BOM present */
    


    while (bytesRead > 0)
    {
        i = 0;
        while (bytesRead > 0 && i < MAXSIZE*7 && nextChar != (TCHAR) '\r')
        {
            buffer[i++] = nextChar;
            ReadFile(btnFile, &nextChar, sizeof(TCHAR), &bytesRead, NULL);
        }
        
        if (bytesRead > 0) ReadFile(btnFile, &nextChar, sizeof(TCHAR), &bytesRead, NULL);  /* read another character to skip line feed */
        buffer[i] = buffer[i+1] = buffer[i+2] = buffer[i+3] = buffer[i+4] = buffer[i+5] = buffer[i+6] = TCHAR(',');
        buffer[i+7] = 0;

        //DEBUG: MessageBox(nppData._nppHandle, buffer, TEXT("Debug: Line Read"), MB_OK);

        if (i > 0 && buffer[0] != (TCHAR) ';')  /* buffer is not empty and does not contain comment */
        {
            for (i = 0, j = 0; j < MAXSIZE-1 && buffer[i] != (TCHAR) ','; i++, j++) g_customMenuStrings[g_customButtonsCount][0][j] = buffer[i];
            g_customMenuStrings[g_customButtonsCount][0][j] = 0;
            // Debugging: Show the first parsed menu string
            //DEBUG: MessageBox(nppData._nppHandle, g_customMenuStrings[g_customButtonsCount][0], TEXT("Parsed Menu String 1"), MB_OK);
            // Debug

            for (i++, j = 0; j < MAXSIZE-1 && buffer[i] != (TCHAR) ','; i++, j++) g_customMenuStrings[g_customButtonsCount][1][j] = buffer[i];
            g_customMenuStrings[g_customButtonsCount][1][j] = 0;
            // Debugging: Show the first parsed menu string
            //DEBUG: MessageBox(nppData._nppHandle, g_customMenuStrings[g_customButtonsCount][1], TEXT("Parsed Menu String 2"), MB_OK);
            // Debug

            for (i++, j = 0; j < MAXSIZE-1 && buffer[i] != (TCHAR) ','; i++, j++) g_customMenuStrings[g_customButtonsCount][2][j] = buffer[i];
            g_customMenuStrings[g_customButtonsCount][2][j] = 0;
            // Debugging: Show the first parsed menu string
            //DEBUG: MessageBox(nppData._nppHandle, g_customMenuStrings[g_customButtonsCount][2], TEXT("Parsed Menu String 3"), MB_OK);
            // Debug

            for (i++, j = 0; j < MAXSIZE-1 && buffer[i] != (TCHAR) ','; i++, j++) g_customMenuStrings[g_customButtonsCount][3][j] = buffer[i];
            g_customMenuStrings[g_customButtonsCount][3][j] = 0;
            // Debugging: Show the first parsed menu string
            //DEBUG: MessageBox(nppData._nppHandle, g_customMenuStrings[g_customButtonsCount][3], TEXT("Parsed Menu String 4"), MB_OK);
            // Debug

            for (i++, j = 0; j < MAXSIZE-1 && buffer[i] != (TCHAR) ','; i++, j++) bmpFileName[j] = buffer[i];
            bmpFileName[j] = 0;
            for (i++, j = 0; j < MAXSIZE-1 && buffer[i] != (TCHAR) ','; i++, j++) icoFileName[j] = buffer[i];
            icoFileName[j] = 0;
            for (i++, j = 0; j < MAXSIZE-1 && buffer[i] != (TCHAR) ','; i++, j++) icodarkFileName[j] = buffer[i];
            icodarkFileName[j] = 0;
            
            //DEBUG: MessageBox(nppData._nppHandle, buffer, TEXT("Debug: Parsed Button Definition"), MB_OK);
            //DEBUG: MessageBox(NULL, bmpFileName, TEXT("Debug: Parsed Bitmap Filename"), MB_OK);
            //DEBUG: MessageBox(NULL, icoFileName, TEXT("Debug: Parsed Icon Filename"), MB_OK);
            //DEBUG: MessageBox(NULL, icodarkFileName, TEXT("Debug: Parsed Dark Icon Filename"), MB_OK);

            bmpError = icoError = icodarkError = false;
            
            if (bmpFileName[0] == (TCHAR) '*') hToolbarBmp = createBitmapForCustomButton(bmpFileName);
            else
            {
                lstrcpy(bmpFilePath, configPath);
                lstrcat(bmpFilePath, TEXT("\\"));
                lstrcat(bmpFilePath, bmpFileName);
                bmpFile = CreateFile(bmpFilePath, GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
                CloseHandle(bmpFile);
                bmpError = (bmpFile == INVALID_HANDLE_VALUE || GetLastError() == ERROR_FILE_NOT_FOUND);
                if (!bmpError) hToolbarBmp = (HBITMAP) LoadImage(NULL, bmpFilePath, IMAGE_BITMAP, 0, 0, (LR_DEFAULTSIZE | LR_LOADTRANSPARENT | LR_LOADMAP3DCOLORS | LR_LOADFROMFILE));
                else hToolbarBmp = (HBITMAP) LoadImage((HINSTANCE) g_hModule, MAKEINTRESOURCE(IDB_CUSTOM_MISSINGFILE), IMAGE_BITMAP, 0, 0, (LR_DEFAULTSIZE | LR_LOADTRANSPARENT | LR_LOADMAP3DCOLORS));
            }
            
            if (icoFileName[0] == (TCHAR) '*') hToolbarIcon = createIconForCustomButton(icoFileName);
            else
            {
                lstrcpy(icoFilePath, configPath);
                lstrcat(icoFilePath, TEXT("\\"));
                lstrcat(icoFilePath, icoFileName);
                icoFile = CreateFile(icoFilePath, GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
                CloseHandle(icoFile);
                icoError = (icoFile == INVALID_HANDLE_VALUE || GetLastError() == ERROR_FILE_NOT_FOUND);
                if (!icoError) hToolbarIcon = (HICON) LoadImage(NULL, icoFilePath, IMAGE_ICON, 0, 0, (LR_DEFAULTSIZE | LR_LOADTRANSPARENT | LR_LOADMAP3DCOLORS | LR_LOADFROMFILE));
                else hToolbarIcon = (HICON) LoadImage((HINSTANCE) g_hModule, MAKEINTRESOURCE(IDI_CUSTOM_MISSINGFILE), IMAGE_ICON, 0, 0, (LR_DEFAULTSIZE | LR_LOADTRANSPARENT | LR_LOADMAP3DCOLORS));
            }
            
            if (icodarkFileName[0] == (TCHAR) '*') hToolbarIconDarkMode = createIconForCustomButton(icodarkFileName);
            else
            {
                lstrcpy(icodarkFilePath, configPath);
                lstrcat(icodarkFilePath, TEXT("\\"));
                lstrcat(icodarkFilePath, icodarkFileName);
                icodarkFile = CreateFile(icodarkFilePath, GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
                CloseHandle(icodarkFile);
                icodarkError = (icodarkFile == INVALID_HANDLE_VALUE || GetLastError() == ERROR_FILE_NOT_FOUND);
                if (!icodarkError) hToolbarIconDarkMode = (HICON) LoadImage(NULL, icodarkFilePath, IMAGE_ICON, 0, 0, (LR_DEFAULTSIZE | LR_LOADTRANSPARENT | LR_LOADMAP3DCOLORS | LR_LOADFROMFILE));
                else
                {
                    if (!icoError) hToolbarIconDarkMode = hToolbarIcon;
                    else hToolbarIconDarkMode = (HICON) LoadImage((HINSTANCE) g_hModule, MAKEINTRESOURCE(IDI_CUSTOM_MISSINGFILE), IMAGE_ICON, 0, 0,(LR_DEFAULTSIZE | LR_LOADTRANSPARENT | LR_LOADMAP3DCOLORS));
                }
            }
            
            if (HIWORD(g_nppVersion) < 8)  /* Notepad++ <= 7.9.5 */
            {
                buttonIcon.hToolbarBmp = hToolbarBmp;
                buttonIcon.hToolbarIcon = hToolbarIcon;
                /* Note: buttonIcon.hToolbarIcon is ignored by Notepad++ <= 7.9.5 */
                SendMessage(nppData._nppHandle, NPPM_ADDTOOLBARICON_DEPRECATED, (WPARAM) (ID_CMD_CUSTOM+g_customButtonsCount), (LPARAM) &buttonIcon);
            }
            else  /* Notepad++ >= 8.0 */
            {
                buttonIconDM.hToolbarBmp = hToolbarBmp;
                buttonIconDM.hToolbarIcon = hToolbarIcon;
                buttonIconDM.hToolbarIconDarkMode = hToolbarIconDarkMode;
                SendMessage(nppData._nppHandle, NPPM_ADDTOOLBARICON_FORDARKMODE, (WPARAM) (ID_CMD_CUSTOM+g_customButtonsCount), (LPARAM) &buttonIconDM);
            }
            
            g_customButtonsCount++;
        }
        
        if (bytesRead > 0) ReadFile(btnFile, &nextChar, sizeof(TCHAR), &bytesRead, NULL);
    }
    
    CloseHandle(btnFile);
}

void addAdditionalButton(int bitmapName, int iconName, int idCmd)
{
    toolbarIcons buttonIcon;
    toolbarIconsWithDarkMode buttonIconDM;
    
    if (HIWORD(g_nppVersion) < 8)  /* Notepad++ <= 7.9.5 */
    {
        buttonIcon.hToolbarBmp = (HBITMAP) LoadImage((HINSTANCE) g_hModule, MAKEINTRESOURCE(bitmapName), IMAGE_BITMAP, 0, 0, (LR_DEFAULTSIZE | LR_LOADTRANSPARENT | LR_LOADMAP3DCOLORS));
        buttonIcon.hToolbarIcon = (HICON) LoadImage((HINSTANCE) g_hModule, MAKEINTRESOURCE(iconName), IMAGE_ICON, 0, 0, (LR_DEFAULTSIZE | LR_LOADTRANSPARENT | LR_LOADMAP3DCOLORS));
        /* Note: buttonIcon.hToolbarIcon is ignored by Notepad++ <= 7.9.5 */
        SendMessage(nppData._nppHandle, NPPM_ADDTOOLBARICON_DEPRECATED, (WPARAM) idCmd, (LPARAM) &buttonIcon);
    }
    else  /* Notepad++ >= 8.0 */
    {
        buttonIconDM.hToolbarBmp = (HBITMAP) LoadImage((HINSTANCE) g_hModule, MAKEINTRESOURCE(bitmapName), IMAGE_BITMAP, 0, 0, (LR_DEFAULTSIZE | LR_LOADTRANSPARENT | LR_LOADMAP3DCOLORS));
        buttonIconDM.hToolbarIcon = (HICON) LoadImage((HINSTANCE) g_hModule, MAKEINTRESOURCE(iconName), IMAGE_ICON, 0, 0, (LR_DEFAULTSIZE | LR_LOADTRANSPARENT | LR_LOADMAP3DCOLORS));
        buttonIconDM.hToolbarIconDarkMode = buttonIconDM.hToolbarIcon;
        SendMessage(nppData._nppHandle, NPPM_ADDTOOLBARICON_FORDARKMODE, (WPARAM) idCmd, (LPARAM) &buttonIconDM);
    }
}

void afterNppReady()
{
    CreateThread(NULL, 0, afterNppReadyDelayed, NULL, 0, NULL);
}

DWORD WINAPI afterNppReadyDelayed(LPVOID lpParam)
{
    HWND rbWindow,tbWindow;
    TCHAR configPath[MAX_PATH];
    TCHAR datFilePath[MAX_PATH];
    HANDLE datFile;
    
    Sleep(10);  /* allow time for other plugins to create additional menu items */
    
    rbWindow = FindWindowEx(nppData._nppHandle, NULL, REBARCLASSNAME, NULL);
    tbWindow = FindWindowEx(rbWindow, NULL, TOOLBARCLASSNAME, NULL);
    
    // Initialize REBARBANDINFO structure size - taking account of Windows and Common Controls versions
    // Constant REBARBANDINFO_V6_SIZE specifies structure size for Common Controls 4.x & 5.x (not 6.x) !
    
    if (!IsWindowsVistaOrGreater() || getCommCtrlMajorVersion() < 6) g_rebarBandInfoSize = REBARBANDINFO_V6_SIZE;  /* Windows pre-Vista or Common Controls pre-6.0 */
    else g_rebarBandInfoSize = sizeof(REBARBANDINFO);
    
    // Subclass the window procedure
    
    g_origWindowProc = (WNDPROC) (LONG_PTR) SetWindowLongPtr(nppData._nppHandle, GWLP_WNDPROC, (LONG_PTR) subclassWindowProc);
    
    // Subclass the rebar procedure
    
    g_origRebarProc = (WNDPROC) (LONG_PTR) SetWindowLongPtr(rbWindow, GWLP_WNDPROC, (LONG_PTR) subclassRebarProc);
    
    // Replace temporary custom command identifiers with actual command identifiers
    // This cannot be done when NPPN_TBMODIFICATION received or immediately after NPPN_READY received,
    // because other plugins may receive these notifications afterwards and create additional menu items
    
    replaceTemporaryCmdIDs();
    
    // Preserve initial toolbar buttons
    
    preserveToolbarButtons();
    
    // Get plugins config directory
    
    SendMessage(nppData._nppHandle, NPPM_GETPLUGINSCONFIGDIR, MAX_PATH, (LPARAM) configPath);
    
    // Initialise .dat file path
    
    lstrcpy(datFilePath, configPath);
    lstrcat(datFilePath, TEXT("\\CustomizeToolbar.dat"));
    
    // Try to open .dat file
    
    datFile = CreateFile(datFilePath, GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    CloseHandle(datFile);
    
    // Reset and save toolbar layout if .dat file does not exist
    
    if (datFile == INVALID_HANDLE_VALUE)
    {
        resetToolbarLayout();
        saveToolbarLayout();
    }
    
    // Restore toolbar layout
    
    restoreToolbarLayout(true);
    updateToolbarState();
    adjustIdealSize();
    
    // Restore toolbar wrap state and display styles
    
    if (g_wrapToolbarState) makeToolbarWrap();
    else makeToolbarOverflow();
    
    return 0;
}

void beforeNppShutdown()
{
    saveToolbarLayout();
}

//
// Menu command functions
//

void customizeToolbar()
{
    HWND rbWindow, tbWindow;
    
    rbWindow = FindWindowEx(nppData._nppHandle, NULL, REBARCLASSNAME, NULL);
    tbWindow = FindWindowEx(rbWindow, NULL, TOOLBARCLASSNAME, NULL);
    
    SendMessage(tbWindow, TB_CUSTOMIZE, (WPARAM) 0, (LPARAM) 0);
}

void customButtons()
{
    TCHAR configPath[MAX_PATH];
    TCHAR btnFilePath[MAX_PATH];
    HANDLE btnFile;
    DWORD bytesWritten;
    const TCHAR *buttonDefs = TEXT("\r\n")
                        TEXT(";EXAMPLES OF CUSTOM BUTTON DEFINITIONS\r\n\r\n")
                        TEXT(";Define custom button for Notepad++ 'Select All' menu command, either using file names:\r\n\r\n")
                        TEXT(";Edit,Select All,,,standard-1.bmp,fluentlight-1.ico,fluentdark-1.ico\r\n\r\n")
                        TEXT(";or using quick codes:\r\n\r\n")
                        TEXT(";Edit,Select All,,,*R:SA,*R:SA\r\n\r\n")
                        TEXT(";Define custom button for Compare plugin 'Settings...' menu command, either using file names:\r\n\r\n")
                        TEXT(";Plugins,Compare,Settings...,,standard-2.bmp,fluentlight-2.ico,fluentdark-2.ico\r\n\r\n")
                        TEXT(";or using quick codes:\r\n\r\n")
                        TEXT(";Plugins,Compare,Settings...,,*G:S,*G:S\r\n\r\n")
                        TEXT(";Redefine existing button for Compare plugin 'Navigation Bar' menu command, either using file names:\r\n\r\n")
                        TEXT(";Plugins,Compare,Navigation Bar,,standard-3.bmp,fluentlight-3.ico,fluentdark-3.ico\r\n\r\n")
                        TEXT(";or using quick codes:\r\n\r\n")
                        TEXT(";Plugins,Compare,Navigation Bar,,*#309030:NB,*#309030:NB\r\n\r\n")
                        TEXT(";With Notepad++ 7.9.5 or earlier, the fluent light and fluent dark fields are ignored.\r\n\r\n")
                        TEXT(";With Notepad++ 8.0 or later, the fluent dark field if omitted defaults to the fluent light field.\r\n");
    
    g_customButtonsState = g_customButtonsState;
    
    SendMessage(nppData._nppHandle, NPPM_SETMENUITEMCHECK, funcItem[2]._cmdID, (LPARAM) g_customButtonsState);
    
    if (g_customButtonsState)
    {
        SendMessage(nppData._nppHandle, NPPM_GETPLUGINSCONFIGDIR, MAX_PATH, (LPARAM) configPath);
        
        lstrcpy(btnFilePath, configPath);
        lstrcat(btnFilePath, TEXT("\\CustomizeToolbar.btn"));
        
        btnFile = CreateFile(btnFilePath, GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
        
        if (btnFile == INVALID_HANDLE_VALUE)
        {
            btnFile = CreateFile(btnFilePath, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
            WriteFile(btnFile, buttonDefs, (DWORD) _tcslen(buttonDefs)*sizeof(TCHAR), &bytesWritten, NULL);
            CloseHandle(btnFile);
        }
        
        MessageBox(nppData._nppHandle, TEXT("Custom Buttons will be enabled on next restart of Notepad++.\n\n"),
                                       TEXT("Customize Toolbar - Custom Buttons - Enabled"), MB_OK | MB_APPLMODAL);
    }
    else
    {
        MessageBox(nppData._nppHandle, TEXT("Custom Buttons will be disabled on next restart of Notepad++.\n\n"),
                                       TEXT("Customize Toolbar - Custom Buttons - Disabled"), MB_OK | MB_APPLMODAL);
    }
    
    saveToolbarLayout();
}

void wrapToolbar()
{
    g_wrapToolbarState = ~g_wrapToolbarState;
    
    SendMessage(nppData._nppHandle, NPPM_SETMENUITEMCHECK, funcItem[3]._cmdID, (LPARAM) g_wrapToolbarState);
    
    // Restore toolbar wrap state and display styles
    
    if (g_wrapToolbarState) makeToolbarWrap();
    else makeToolbarOverflow();
    
    saveToolbarLayout();
}

void helpOverview()
{
    MessageBox(nppData._nppHandle, TEXT("Customize Toolbar Plugin\n\n")
                                   TEXT("Version: 5.3    -    © 2011-2021 DW-dev    -    E-mail: dw-dev@gmx.com\n\n")
                                   TEXT("Version: 5.3.1    -    © 2024+   QGtKMlLz    -    E-mail: 3m33dkojb@mozmail.com\n\n")
                                   TEXT("This plugin allows the Notepad++ toolbar to be fully customised by the user, and includes twenty-six additional buttons for frequently used menu commands.\n\n")
                                   TEXT("All buttons on the toolbar can be customized, whether Notepad++ built-in buttons, the additional buttons, or buttons belonging to other plugins.\n\n")
                                   TEXT("When this plugin is first installed, the additional buttons are not shown on the toolbar, but are available in the Customize Toolbar dialog box.\n\n")
                                   TEXT("The toolbar is customized using the Customize Toolbar dialog box, which can be opened by clicking on the Customize Toolbar... menu item, or ")
                                   TEXT("by clicking on the Customize Toolbar... toolbar button, or by double-clicking on empty space on the toolbar.\n\n")
                                   TEXT("Alternatively, the toolbar can be customized by holding down the Shift key and dragging a button along the toolbar or off the toolbar.\n\n")
                                   TEXT("It is recommended to customize the toolbar when Standard Icons are selected in Notepad++ preferences, so that buttons belonging to other plugins are visible.\n\n")
                                   TEXT("Custom buttons for Notepad++ or plugin menu commands can be defined using a configuration file, and there is a menu option to enable/disable this feature.\n\n")
                                   TEXT("An overflow chevron is shown if there are too many buttons to fit on the toolbar. Alternatively, there is a menu option to wrap the toolbar over several rows.\n\n")
                                   TEXT("There is a menu option to show the resources (toolbar buttons and plugin menu commands) that are currently being used.\n\n"),
                                   TEXT("Customize Toolbar - Help - Overview"), MB_OK | MB_APPLMODAL);
}

void helpCustomButtons()
{
    MessageBox(nppData._nppHandle, TEXT("Custom buttons are defined using a configuration file (CustomizeToolbar.btn) located in the Notepad++ configuration sub-folder (...\\plugins\\config).\n\n")
                                   TEXT("When the Custom Buttons feature is enabled, if the .btn configuration file did not previously exist, it is created and contains examples of custom button definitions.\n\n")
#ifdef _UNICODE
                                   TEXT("The .btn configuration file must employ Unicode UTF-16 Little Endian encoding, with an optional Byte Order Mark (BOM) at the start of file, and CR-LF line breaks.  ")
                                   TEXT("When creating this file with Notepad++, set Encoding to UCS-2 Little Endian.\n\n")
#else
                                   TEXT("The .btn configuration file must employ ANSI encoding and CR-LF line breaks.  ")
                                   TEXT("When creating this file with Notepad++, set Encoding to ANSI.\n\n")
#endif
                                   TEXT("Each line in the .btn configuration file can be either a custom button definition or a comment starting with a semicolon.\n\n")
                                   TEXT("Each custom button definition comprises seven comma separated fields (four menu strings, an optional .bmp file name for Standard icons, ")
                                   TEXT("and two optional .ico file names for Fluent icons in light and dark modes).\n\n")
                                   TEXT("If the menu strings correspond to a Notepad++ built-in button or plugin button, the custom button will replace the Notepad++ built-in button or plugin button.\n\n")
                                   TEXT("If the menu strings do not correspond to a Notepad++ built-in button or plugin button, then an error symbol (exclamation mark) is displayed.\n\n")
                                   TEXT("If the .bmp or .ico file names are present, the files must be located in the Notepad++ configuration sub-folder (...\\plugins\\config).\n\n")
                                   TEXT("If the .bmp or light mode .ico file name is omitted, or if the file does not exist, then a warning symbol (question mark) is displayed.\n\n")
                                   TEXT("If the dark mode .ico file name is omitted, or if the file does not exist, then if present the light mode .ico file name is used instead.\n\n")
                                   TEXT("Each .bmp file must be an image of 16x16 pixels with a bit depth of 8-bits. Any pixels with the same colour as the bottom left pixel will appear transparent.\n\n")
                                   TEXT("Each .ico file must be an icon containing an image of 32x32 pixels with a bit depth of 32-bits (RGB+alpha).\n\n")
                                   TEXT("Quick codes can be used instead of file names. A quick code comprises:\nan asterisk, followed by either a color code letter (S: slate grey, R: red,\n")
                                   TEXT("G: green, B: blue, C: cyan, M: magenta, Y: yellow) or a hex color value\n(e.g. #4488CC), followed by a colon, followed by a label (1 or 2 letters).\n\n")
                                   TEXT("To create a red button with label 'LA', use: *R:LA or *#FF0000:LA.\n\n"),
                                   TEXT("Customize Toolbar - Help - Custom Buttons"), MB_OK | MB_APPLMODAL);
}

void resourceUsage()
{
    TCHAR buffer[100];
    int commands, maxcommands;
    
    commands = funcItem[8]._cmdID-ID_PLUGINS_CMD+1;
    maxcommands = g_id_plugins_cmd_limit-ID_PLUGINS_CMD+1;
    
    _stprintf_s(buffer, 100, TEXT("Total Buttons:  %i / 300\n\nCustom Buttons:  %i / 100\n\nPlugin Menu Commands:  %i / %i\n"), g_buttonsAvailable, g_customButtonsCount, commands, maxcommands);
    
    MessageBox(nppData._nppHandle, buffer, TEXT("Customize Toolbar - Resource Usage"), MB_OK | MB_APPLMODAL);
}

//
// Subclass window procedure and rebar procedure
//

LRESULT APIENTRY subclassWindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    HWND rbWindow,tbWindow;
    NMTOOLBAR *lpNmToolbar = (LPNMTOOLBAR) lParam;
    NMREBARCHEVRON *lpNmRebarChevron = (LPNMREBARCHEVRON) lParam;
    
    rbWindow = FindWindowEx(nppData._nppHandle,NULL,REBARCLASSNAME,NULL);
    tbWindow = FindWindowEx(rbWindow,NULL,TOOLBARCLASSNAME,NULL);
    
    if (uMsg == WM_GETDLGCODE)
        return DLGC_WANTALLKEYS;
    
    // Handle toolbar customization
    
    if (uMsg == WM_NOTIFY)
    {
        switch (lpNmToolbar->hdr.code)
        {
            case TBN_INITCUSTOMIZE:
                
                // Hide help button
                
                return TBNRF_HIDEHELP;
                break;
                
            case TBN_GETBUTTONINFO:
                
                if (lpNmToolbar->iItem < g_buttonsAvailable)
                {
                    // Pass the next button
                    
                    lpNmToolbar->tbButton = g_tbButtons[lpNmToolbar->iItem];
                    return TRUE;
                }
                else
                {
                    // No more buttons
                    
                    return FALSE;
                }
                break;
                
            case TBN_QUERYINSERT:
                
                // Returning FALSE causes customize dialog box to not appear.
                
                // Restore toolbar wrap state and display styles
                
                if (g_wrapToolbarState) makeToolbarWrap();
                else makeToolbarOverflow();
                
                return TRUE;
                break;
                
            case TBN_QUERYDELETE:
                
                // Returning FALSE causes customize dialog box to not appear.
                
                return TRUE;
                break;
                
            case TBN_RESET:
                
                // Reset toolbar to un-customised layout
                
                resetToolbarLayout();
                
                return TRUE;
                break;
                
            case TBN_ENDADJUST:
                
                // Restore toolbar wrap state and display styles
                
                if (g_wrapToolbarState) makeToolbarWrap();
                else makeToolbarOverflow();
                
                // Save customized toolbar
                
                updateToolbarState();
                adjustIdealSize();
                saveToolbarLayout();
                
                return 0;  /* not used */
                break;
                
            case RBN_ENDDRAG:
                
                // Restore toolbar wrap state and display styles
                
                if (g_wrapToolbarState) makeToolbarWrap();
                else makeToolbarOverflow();
                
                break;
                
            case RBN_CHEVRONPUSHED:
                
                // Display popup menu of overflow buttons
                
                displayOverflowMenu(lpNmRebarChevron);
                
                return 0;  /* not used */
                break;
                
            case RBN_CHILDSIZE:
                
                // Toolbar window resized - not used at present
                
                break;
                
            case NM_CLICK:
                
                // Restore toolbar wrap state and display styles
                
                if (g_wrapToolbarState) makeToolbarWrap();
                else makeToolbarOverflow();
                
                // Update toolbar button states after toolbar clicked
                
                CreateThread(NULL, 0, handleButtonStates, NULL, 0, NULL);
                
                break;
        }
    }
    
    // Handle window re-size
    
    if (uMsg == WM_SIZE)
    {
        CreateThread(NULL, 0, handleWindowResize, NULL, 0, NULL);
    }
    
    // Handle update of button states after menu command selected
    
    if (uMsg == WM_UNINITMENUPOPUP)
    {
        CreateThread(NULL, 0, handleButtonStates, NULL, 0, NULL);
    }
    
    return CallWindowProc(g_origWindowProc, hwnd, uMsg, wParam, lParam);
}

LRESULT APIENTRY subclassRebarProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    REBARBANDINFO  *lpRebarBandInfo = (LPREBARBANDINFO) lParam;
    
    if (uMsg == WM_GETDLGCODE)
        return DLGC_WANTALLKEYS;
    
    if (uMsg == RB_SETBANDINFO && lpRebarBandInfo->fMask == 0x0270)  /* toolbar has been reset and icons changed by Notepad++ */
    {
        CreateThread(NULL, 0, handleChangedIcons, NULL, 0, NULL);
    }
    
    return CallWindowProc(g_origRebarProc, hwnd, uMsg, wParam, lParam);
}

//
// Handle window resize, changed icons and button states thread delay functions
//

DWORD WINAPI handleWindowResize(LPVOID lpParam)
{
    Sleep(10);
    
    // Restore toolbar wrap state and display styles
    
    if (g_wrapToolbarState) makeToolbarWrap();
    else makeToolbarOverflow();
    
    return 0;
}

DWORD WINAPI handleChangedIcons(LPVOID lpParam)
{
    Sleep(10);
    
    // Replace temporary custom command identifiers with actual command identifiers
    
    replaceTemporaryCmdIDs();
    
    // Preserve initial toolbar buttons
    
    preserveToolbarButtons();
    
    // Restore toolbar layout
    
    restoreToolbarLayout(false);
    updateToolbarState();
    adjustIdealSize();
    
    // Restore toolbar wrap state and display styles
    
    if (g_wrapToolbarState) makeToolbarWrap();
    else makeToolbarOverflow();
    
    return 0;
}

DWORD WINAPI handleButtonStates(LPVOID lpParam)
{
    Sleep(10);
    
    // Update toolbar button states after toolbar is clicked or menu item is seleceted
    
    updateToolbarState();
    
    return 0;
}

//
// Replace temporary custom command identifiers function
//

void replaceTemporaryCmdIDs()
{
    HWND rbWindow, tbWindow;
    int i, j, btn, idCmd;
    
    rbWindow = FindWindowEx(nppData._nppHandle, NULL, REBARCLASSNAME, NULL);
    tbWindow = FindWindowEx(rbWindow, NULL, TOOLBARCLASSNAME, NULL);
    
    for (btn = 0; btn < g_customButtonsCount; btn++)
    {
        idCmd = findCmdIDForMenuStrings(g_hMainMenu, g_customMenuStrings[btn][0], g_customMenuStrings[btn][1], g_customMenuStrings[btn][2], g_customMenuStrings[btn][3]);
        
        if (idCmd != -1)  /* if menu strings found */
        {
            // remove built-in or plugin button (if any) with this command identifier
            
            i = (DWORD) SendMessage(tbWindow, TB_COMMANDTOINDEX, (WPARAM) (idCmd), (LPARAM) 0);
            if (i != -1) SendMessage(tbWindow, TB_DELETEBUTTON, (WPARAM) i, (LPARAM) 0);
            
            // replace temporary custom button command identifier
            
            j = (DWORD) SendMessage(tbWindow, TB_COMMANDTOINDEX, (WPARAM) (ID_CMD_CUSTOM+btn), (LPARAM) 0);
            SendMessage(tbWindow, TB_SETCMDID, (WPARAM) j, (LPARAM) idCmd);
        }
    }
}

//
// Preserve initial toolbar buttons and update toolbar button state functions
//

void preserveToolbarButtons()
{
    HWND rbWindow, tbWindow;
    TCHAR buffer[MAXSIZE*4+50];  /* menu string or error message with 4 menu strings */
    HIMAGELIST hImageList;
    HICON hIcon;
    MENUITEMINFO menuItemInfo;
    int i,scriptCount;
    
    rbWindow = FindWindowEx(nppData._nppHandle, NULL, REBARCLASSNAME, NULL);
    tbWindow = FindWindowEx(rbWindow, NULL, TOOLBARCLASSNAME, NULL);
    
    // Preserve startup toolbar button count (for reset and save/restore)
    
    g_buttonsAvailable = (int) SendMessage(tbWindow, TB_BUTTONCOUNT, (WPARAM) 0, (LPARAM) 0);
    
    // Preserve startup toolbar button information (for reset and save/restore)
    
    for (i = 0; i < g_buttonsAvailable; i++)
    {
        SendMessage(tbWindow, TB_GETBUTTON, (WPARAM) i, (LPARAM)(LPTBBUTTON) &g_tbButtons[i]);
        
        if (g_tbButtons[i].idCommand < ID_CMD_CUSTOM || g_tbButtons[i].idCommand > ID_CMD_CUSTOM_LIMIT)
        {
            GetMenuString(g_hMainMenu, g_tbButtons[i].idCommand, buffer, MAXSIZE, MF_BYCOMMAND);
            stripMenuString(buffer);
        }
        else
        {
            hImageList = (HIMAGELIST) SendMessage(tbWindow, TB_GETIMAGELIST, (WPARAM) 0, (LPARAM) 0);
            hIcon = (HICON) LoadImage((HINSTANCE) g_hModule, MAKEINTRESOURCE(IDI_CUSTOM_FAILEDMATCH), IMAGE_ICON, 0, 0,(LR_DEFAULTSIZE | LR_LOADTRANSPARENT | LR_LOADMAP3DCOLORS));
            ImageList_ReplaceIcon(hImageList, g_tbButtons[i].iBitmap, hIcon);
            SendMessage(tbWindow, TB_SETIMAGELIST, (WPARAM) 0, (LPARAM) hImageList);

            hImageList = (HIMAGELIST) SendMessage(tbWindow, TB_GETDISABLEDIMAGELIST, (WPARAM) 0, (LPARAM) 0);
            hIcon = (HICON) LoadImage((HINSTANCE) g_hModule, MAKEINTRESOURCE(IDI_CUSTOM_FAILEDMATCH), IMAGE_ICON, 0, 0,(LR_DEFAULTSIZE | LR_LOADTRANSPARENT | LR_LOADMAP3DCOLORS));
            ImageList_ReplaceIcon(hImageList, g_tbButtons[i].iBitmap, hIcon);
            SendMessage(tbWindow, TB_SETDISABLEDIMAGELIST, (WPARAM) 0, (LPARAM) hImageList);
            
            lstrcpy(buffer, TEXT("Custom Button Error: "));
            lstrcat(buffer, g_customMenuStrings[g_tbButtons[i].idCommand-ID_CMD_CUSTOM][0]);
            lstrcat(buffer, TEXT(","));
            lstrcat(buffer, g_customMenuStrings[g_tbButtons[i].idCommand-ID_CMD_CUSTOM][1]);
            lstrcat(buffer, TEXT(","));
            lstrcat(buffer, g_customMenuStrings[g_tbButtons[i].idCommand-ID_CMD_CUSTOM][2]);
            lstrcat(buffer, TEXT(","));
            lstrcat(buffer, g_customMenuStrings[g_tbButtons[i].idCommand-ID_CMD_CUSTOM][3]);
        }
        
        buffer[_tcslen(buffer)+1] = 0;  /* TB_ADDSTRING requires two null characters */
        
        g_tbButtons[i].iString = SendMessage(tbWindow, TB_ADDSTRING, 0, (LPARAM) buffer);
        
        g_tbButtons[i].fsState = TBSTATE_ENABLED;
    }
    
    // WebEdit Workaround - remove "WebEdit - " from menu strings - as WebEdit will do when it initialises
    
    for (i = 0; i < g_buttonsAvailable; i++)
    {
        if (g_tbButtons[i].idCommand >= ID_PLUGINS_CMD && g_tbButtons[i].idCommand <= g_id_plugins_cmd_limit)  /* plugin command (with menu item) */
        {
            GetMenuString(g_hMainMenu, g_tbButtons[i].idCommand, buffer, MAXSIZE, MF_BYCOMMAND);
            if (_tcsncmp(buffer, TEXT("WebEdit - "), 10) == 0)
            {
                menuItemInfo.cbSize = sizeof(MENUITEMINFO);
                menuItemInfo.fMask = MIIM_STRING;
                menuItemInfo.dwTypeData = buffer+10;
                SetMenuItemInfo(g_hMainMenu, g_tbButtons[i].idCommand, false, &menuItemInfo);
            }
        }
    }
    
    // Python Script Workaround - add button strings - since Python Script button commands are not on menu
    
    scriptCount = 1;
    for (i = 0; i < g_buttonsAvailable; i++)
    {
        if (g_tbButtons[i].idCommand >= ID_PLUGINS_CMD_DYNAMIC && g_tbButtons[i].idCommand <= ID_PLUGINS_CMD_DYNAMIC_LIMIT)  /* plugin command (without menu item) (from NPPM_ALLOCATECMDID) */
        {
            GetMenuString(g_hMainMenu, g_tbButtons[i].idCommand, buffer, MAXSIZE, MF_BYCOMMAND);
            if (buffer[0] == 0)
            {
                lstrcpy(buffer, TEXT("Python Script "));
                _itot_s(scriptCount++, buffer+MAXSIZE, MAXSIZE, 10);
                lstrcat(buffer, buffer+MAXSIZE);
                
                buffer[_tcslen(buffer)+1] = 0;  /* TB_ADDSTRING requires two null characters */
                
                g_tbButtons[i].iString = SendMessage(tbWindow, TB_ADDSTRING, 0, (LPARAM) buffer);
            }
        }
    }
}

void updateToolbarState()
{
    HWND rbWindow, tbWindow;
    TBBUTTON tbButton;
    int i, buttonsOnToolbar, menuState, tbState;
    
    rbWindow = FindWindowEx(nppData._nppHandle, NULL, REBARCLASSNAME, NULL);
    tbWindow = FindWindowEx(rbWindow, NULL, TOOLBARCLASSNAME, NULL);
    
    buttonsOnToolbar = (int) SendMessage(tbWindow, TB_BUTTONCOUNT, (WPARAM) 0, (LPARAM) 0);
    
    for (i = 0; i < buttonsOnToolbar; i++)
    {
        SendMessage(tbWindow, TB_GETBUTTON, (WPARAM) i, (LPARAM)(LPTBBUTTON) &tbButton);
        menuState = GetMenuState(g_hMainMenu, tbButton.idCommand, MF_BYCOMMAND);
        if (menuState == -1) menuState = 0;  /* no menu associated with button (e.g. Python Script command button) */
        tbState = 0;
        if (tbButton.idCommand < ID_CMD_CUSTOM || tbButton.idCommand > ID_CMD_CUSTOM_LIMIT)
        {
            if (menuState & MF_CHECKED) tbState |= TBSTATE_CHECKED;
            if (!(menuState & (MF_DISABLED | MF_GRAYED))) tbState |= TBSTATE_ENABLED;
        }
        SendMessage(tbWindow, TB_SETSTATE, (WPARAM) tbButton.idCommand, (LPARAM) MAKELPARAM(tbState, 0));
    }
}

//
// Reset, save and restore toolbar layout functions
//

void resetToolbarLayout()
{
    HWND rbWindow, tbWindow;
    int i, idCmd;
    
    rbWindow = FindWindowEx(nppData._nppHandle, NULL, REBARCLASSNAME, NULL);
    tbWindow = FindWindowEx(rbWindow, NULL, TOOLBARCLASSNAME, NULL);
    
    // Remove all buttons from toolbar
    
    i = (int) SendMessage(tbWindow, TB_BUTTONCOUNT, (WPARAM) 0, (LPARAM) 0);
    for (i = i - 1; i >= 0; i--)
    {
        SendMessage(tbWindow, TB_DELETEBUTTON, (WPARAM) i, (LPARAM) 0);
    }
    
    // Add the buttons that were preserved
    
    for (i = 0; i < g_buttonsAvailable; i++)
    {
        idCmd = g_tbButtons[i].idCommand;
        
        if (idCmd == IDM_FILE_CLOSEALL_BUT_CURRENT ||
            idCmd == IDM_EDIT_DELETE || idCmd == IDM_EDIT_RMV_TAB || idCmd == IDM_EDIT_INS_TAB || idCmd == IDM_EDIT_DUP_LINE ||
            idCmd == IDM_EDIT_BLOCK_COMMENT_SET || idCmd == IDM_EDIT_BLOCK_UNCOMMENT || idCmd == IDM_EDIT_AUTOCOMPLETE_CURRENTFILE ||
            idCmd == IDM_EDIT_TRIMTRAILING || idCmd == IDM_EDIT_TAB2SW || idCmd == IDM_EDIT_SW2TAB_ALL ||
            idCmd == IDM_SEARCH_FINDINFILES || idCmd == IDM_SEARCH_FINDPREV || idCmd == IDM_SEARCH_FINDNEXT ||
            idCmd == IDM_SEARCH_FINDINCREMENT || idCmd == IDM_FOCUS_ON_FOUND_RESULTS || idCmd == IDM_SEARCH_GOTOLINE ||
            idCmd == IDM_SEARCH_PREV_BOOKMARK || idCmd == IDM_SEARCH_NEXT_BOOKMARK || idCmd == IDM_SEARCH_CLEAR_BOOKMARKS ||
            idCmd == IDM_VIEW_ZOOMRESTORE || idCmd == IDM_VIEW_GOTO_ANOTHER_VIEW || idCmd == IDM_VIEW_CLONE_TO_ANOTHER_VIEW ||
            idCmd == IDM_VIEW_HIDELINES || idCmd == IDM_VIEW_TOGGLE_FOLDALL || idCmd == IDM_VIEW_TOGGLE_UNFOLDALL) ;  /* do nothing */
        else SendMessage(tbWindow, TB_ADDBUTTONS, (WPARAM)(UINT) 1, (LPARAM)(LPTBBUTTON) &g_tbButtons[i]);
    }
    
    // Without this added buttons are not displayed !!
    
    SendMessage(tbWindow, TB_SETMAXTEXTROWS, (WPARAM) 0, (LPARAM) 0);
}

void saveToolbarLayout()
{
    HWND rbWindow, tbWindow;
    TCHAR configPath[MAX_PATH];
    TCHAR datFilePath[MAX_PATH];
    HANDLE datFile;
    DWORD bytesWritten;
    DWORD dword;
    TBBUTTON tbButton;
    int i, j, buttonsOnToolbar;
    
    rbWindow = FindWindowEx(nppData._nppHandle, NULL, REBARCLASSNAME, NULL);
    tbWindow = FindWindowEx(rbWindow, NULL, TOOLBARCLASSNAME, NULL);
    
    // Get plugins config directory
    
    SendMessage(nppData._nppHandle, NPPM_GETPLUGINSCONFIGDIR, MAX_PATH, (LPARAM) configPath);
    
    // Initialise .dat file path
    
    lstrcpy(datFilePath, configPath);
    lstrcat(datFilePath, TEXT("\\CustomizeToolbar.dat"));
    
    // Create and open .dat file
    
    datFile = CreateFile(datFilePath, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    
    // Write custom buttons menu item state
    
    WriteFile(datFile, &g_customButtonsState, sizeof(int), &bytesWritten, NULL);
    
    // Write wrap toolbar menu item state
    
    WriteFile(datFile, &g_wrapToolbarState, sizeof(int), &bytesWritten, NULL);
    
    // Write count of buttons on toolbar (currently)
    
    buttonsOnToolbar = (int) SendMessage(tbWindow, TB_BUTTONCOUNT, (WPARAM) 0, (LPARAM) 0);
    WriteFile(datFile, &buttonsOnToolbar, sizeof(int), &bytesWritten, NULL);
    
    // Write count of all buttons available (at startup)
    
    WriteFile(datFile, &g_buttonsAvailable, sizeof(int), &bytesWritten, NULL);
    
    // Write entry for each button on toolbar (currently)
    
    for (i = 0; i < buttonsOnToolbar; i++)
    {
        SendMessage(tbWindow, TB_GETBUTTON, (WPARAM) i, (LPARAM)(LPTBBUTTON) &tbButton);
        if (tbButton.idCommand >= ID_PLUGINS_CMD && tbButton.idCommand <= g_id_plugins_cmd_limit)  /* plugin command (with menu item) */
        {
            dword = calcPluginButtonMenuHash(tbButton);
            WriteFile(datFile, &dword, sizeof(DWORD), &bytesWritten, NULL);
        }
        else if (tbButton.idCommand >= ID_PLUGINS_CMD_DYNAMIC && tbButton.idCommand <= ID_PLUGINS_CMD_DYNAMIC_LIMIT)  /* plugin command (without menu item) (from NPPM_ALLOCATECMDID) */
        {
            dword = calcButtonStringHash(tbButton);
            WriteFile(datFile, &dword, sizeof(DWORD), &bytesWritten, NULL);
        }
        else if (tbButton.idCommand >= ID_CMD_CUSTOM && tbButton.idCommand <= ID_CMD_CUSTOM_LIMIT)  /* custom command (menu strings not found) */
        {
            dword = calcButtonStringHash(tbButton);
            WriteFile(datFile, &dword, sizeof(DWORD), &bytesWritten, NULL);
        }
        else  /* built-in command or separator */
        {
            dword = tbButton.idCommand;
            WriteFile(datFile, &dword, sizeof(DWORD), &bytesWritten, NULL);
        }
    }
    
    // Write entry for each button available (at startup)
    
    for (j = 0; j < g_buttonsAvailable; j++)
    {
        if (g_tbButtons[j].idCommand >= ID_PLUGINS_CMD && g_tbButtons[j].idCommand <= g_id_plugins_cmd_limit)  /* plugin command (with menu item) */
        {
            dword = calcPluginButtonMenuHash(g_tbButtons[j]);
            WriteFile(datFile, &dword, sizeof(DWORD), &bytesWritten, NULL);
        }
        else if (g_tbButtons[j].idCommand >= ID_PLUGINS_CMD_DYNAMIC && g_tbButtons[j].idCommand <= ID_PLUGINS_CMD_DYNAMIC_LIMIT)  /* plugin command (without menu item) (from NPPM_ALLOCATECMDID) */
        {
            dword = calcButtonStringHash(g_tbButtons[j]);
            WriteFile(datFile, &dword, sizeof(DWORD), &bytesWritten, NULL);
        }
        else if (g_tbButtons[j].idCommand >= ID_CMD_CUSTOM && g_tbButtons[j].idCommand <= ID_CMD_CUSTOM_LIMIT)  /* custom command (menu strings not found) */
        {
            dword = calcButtonStringHash(g_tbButtons[j]);
            WriteFile(datFile, &dword, sizeof(DWORD), &bytesWritten, NULL);
        }
        else  /* built-in command or separator */
        {
            dword = g_tbButtons[j].idCommand;
            WriteFile(datFile, &dword, sizeof(DWORD), &bytesWritten, NULL);
        }
    }
    
    // Close .dat file
    
    CloseHandle(datFile);
}

void restoreToolbarLayout(bool menuStates)
{
    HWND rbWindow, tbWindow;
    TCHAR configPath[MAX_PATH];
    TCHAR datFilePath[MAX_PATH];
    HANDLE datFile;
    DWORD bytesRead;
    DWORD dword;
    int i, j, buttonsAvailable, buttonsOnToolbar;
    
    rbWindow = FindWindowEx(nppData._nppHandle, NULL, REBARCLASSNAME, NULL);
    tbWindow = FindWindowEx(rbWindow, NULL, TOOLBARCLASSNAME, NULL);
    
    // Get plugins config directory
    
    SendMessage(nppData._nppHandle, NPPM_GETPLUGINSCONFIGDIR, MAX_PATH, (LPARAM) configPath);
    
    // Initialise .dat file path
    
    lstrcpy(datFilePath, configPath);
    lstrcat(datFilePath, TEXT("\\CustomizeToolbar.dat"));
    
    // Open .dat file
    
    datFile = CreateFile(datFilePath, GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    
    // Read custom buttons menu item state
    
    ReadFile(datFile, &i, sizeof(int), &bytesRead, NULL);
    if (menuStates)
    {
        g_customButtonsState = i;
        SendMessage(nppData._nppHandle, NPPM_SETMENUITEMCHECK, funcItem[2]._cmdID, (LPARAM) g_customButtonsState);
    }
    
    // Read wrap toolbar menu item state
    
    ReadFile(datFile, &i, sizeof(int), &bytesRead, NULL);
    if (menuStates)
    {
        g_wrapToolbarState = i;
        SendMessage(nppData._nppHandle, NPPM_SETMENUITEMCHECK, funcItem[3]._cmdID, (LPARAM) g_wrapToolbarState);
    }
    
    // Remove all buttons from toolbar
    
    i = (int) SendMessage(tbWindow, TB_BUTTONCOUNT, (WPARAM) 0, (LPARAM) 0);
    for (i = i - 1; i >= 0; i--)
    {
        SendMessage(tbWindow, TB_DELETEBUTTON, (WPARAM) i, (LPARAM) 0);
    }
    
    // Read count of buttons on toolbar (in last session)
    
    ReadFile(datFile, &buttonsOnToolbar, sizeof(int), &bytesRead, NULL);
    
    // Read count of all buttons available (in last session)
    
    ReadFile(datFile, &buttonsAvailable, sizeof(int), &bytesRead, NULL);
    
    // Read entry for each button on toolbar (in last session)
    
    for (i = 0; i < buttonsOnToolbar; i++)
    {
        ReadFile(datFile, &dword, sizeof(DWORD), &bytesRead, NULL);
        if (dword & HASHFLAG)  /* plugin command */
        {
            for (j = 0; j < g_buttonsAvailable; j++)
            {
                if (g_tbButtons[j].idCommand >= ID_PLUGINS_CMD && g_tbButtons[j].idCommand <= g_id_plugins_cmd_limit)  /* plugin command (with menu item) */
                {
                    if (dword == calcPluginButtonMenuHash(g_tbButtons[j]))
                    {
                        SendMessage(tbWindow, TB_ADDBUTTONS, (WPARAM)(UINT) 1, (LPARAM)(LPTBBUTTON) &g_tbButtons[j]);
                        break;
                    }
                }
                else if (g_tbButtons[j].idCommand >= ID_PLUGINS_CMD_DYNAMIC && g_tbButtons[j].idCommand <= ID_PLUGINS_CMD_DYNAMIC_LIMIT)  /* plugin command (without menu item) (from NPPM_ALLOCATECMDID) */
                {
                    if (dword == calcButtonStringHash(g_tbButtons[j]))
                    {
                        SendMessage(tbWindow, TB_ADDBUTTONS, (WPARAM)(UINT) 1, (LPARAM)(LPTBBUTTON) &g_tbButtons[j]);
                        break;
                    }
                }
                else if (g_tbButtons[j].idCommand >= ID_CMD_CUSTOM && g_tbButtons[j].idCommand <= ID_CMD_CUSTOM_LIMIT)  /* plugin command (menu strings not found) */
                {
                    if (dword == calcButtonStringHash(g_tbButtons[j]))
                    {
                        SendMessage(tbWindow, TB_ADDBUTTONS, (WPARAM)(UINT) 1, (LPARAM)(LPTBBUTTON) &g_tbButtons[j]);
                        break;
                    }
                }
            }
        }
        else  /* built-in command or separator */
        {
            for (j = 0; j < g_buttonsAvailable; j++)
            {
                if (dword == g_tbButtons[j].idCommand)
                {
                    SendMessage(tbWindow, TB_ADDBUTTONS, (WPARAM)(UINT) 1, (LPARAM)(LPTBBUTTON) &g_tbButtons[j]);
                    break;  /* to avoid multiple separators being restored */
                }
            }
        }
    }
    
    // Read entry for each button available (in last session)
    
    for (j = 0; j < g_buttonsAvailable; j++)
    {
        g_tbButtons[j].dwData = 1;
    }
    
    for (i = 0; i < buttonsAvailable; i++)
    {
        ReadFile(datFile, &dword, sizeof(DWORD), &bytesRead, NULL);
        if (dword & HASHFLAG)  /* plugin command */
        {
            for (j = 0; j < g_buttonsAvailable; j++)
            {
                if (g_tbButtons[j].idCommand >= ID_PLUGINS_CMD && g_tbButtons[j].idCommand <= g_id_plugins_cmd_limit)  /* plugin command (with menu item) */
                {
                    if (dword == calcPluginButtonMenuHash(g_tbButtons[j]))
                    {
                        g_tbButtons[j].dwData = 0;
                        break;
                    }
                }
                else if (g_tbButtons[j].idCommand >= ID_PLUGINS_CMD_DYNAMIC && g_tbButtons[j].idCommand <= ID_PLUGINS_CMD_DYNAMIC_LIMIT)  /* plugin command (without menu item) (from NPPM_ALLOCATECMDID) */
                {
                    if (dword == calcButtonStringHash(g_tbButtons[j]))
                    {
                        g_tbButtons[j].dwData = 0;
                        break;
                    }
                }
                else if (g_tbButtons[j].idCommand >= ID_CMD_CUSTOM && g_tbButtons[j].idCommand <= ID_CMD_CUSTOM_LIMIT)  /* plugin command (menu strings not found) */
                {
                    if (dword == calcButtonStringHash(g_tbButtons[j]))
                    {
                        g_tbButtons[j].dwData = 0;
                        break;
                    }
                }
            }
        }
        else  /* built-in command or separator */
        {
            for (j = 0; j < g_buttonsAvailable; j++)
            {
                if (dword == g_tbButtons[j].idCommand)
                {
                    g_tbButtons[j].dwData = 0;
                    //break;   /* no break to ensure all separators have dwData flag cleared */
                }
            }
        }
    }
    
    for (j = 0; j < g_buttonsAvailable; j++)
    {
        if (g_tbButtons[j].dwData == 1)
        {
            g_tbButtons[j].dwData = 0;
            SendMessage(tbWindow, TB_ADDBUTTONS, (WPARAM)(UINT) 1, (LPARAM)(LPTBBUTTON) &g_tbButtons[j]);
        }
    }
    
    // Without this added buttons are not displayed !!
    
    SendMessage(tbWindow, TB_SETMAXTEXTROWS, (WPARAM) 0, (LPARAM) 0);
    
    // Close .dat file
    
    CloseHandle(datFile);
}

//
// Toolbar wrap/overflow, ideal size and overflow menu functions
//

void makeToolbarWrap()
{
    HWND rbWindow, tbWindow;
    LONG_PTR style;
    REBARBANDINFO rebarBandInfo;
    RECT buttonRect;
    DWORD padding;
    int buttonsOnToolbar;
    
    rbWindow = FindWindowEx(nppData._nppHandle, NULL, REBARCLASSNAME, NULL);
    tbWindow = FindWindowEx(rbWindow, NULL, TOOLBARCLASSNAME, NULL);
    
    style = GetWindowLongPtr(tbWindow, GWL_STYLE);
    style |= TBSTYLE_WRAPABLE;
    style |= TBSTYLE_TRANSPARENT;
    style |= CCS_ADJUSTABLE;
    SetWindowLongPtr(tbWindow, GWL_STYLE, style);
    
    buttonsOnToolbar = (int) SendMessage(tbWindow, TB_BUTTONCOUNT, (WPARAM) 0, (LPARAM) 0);
    
    SendMessage(tbWindow, TB_GETITEMRECT, (WPARAM) (buttonsOnToolbar-1), (LPARAM)(LPTBBUTTON) &buttonRect);
    padding = (DWORD) SendMessage(tbWindow, TB_GETPADDING, 0, 0);
    
    rebarBandInfo.cbSize = g_rebarBandInfoSize;
    rebarBandInfo.fMask = RBBIM_STYLE | RBBIM_CHILDSIZE;
    SendMessage(rbWindow, RB_GETBANDINFO, (WPARAM) 0, (LPARAM) &rebarBandInfo);
    
    rebarBandInfo.cbSize = g_rebarBandInfoSize;
    rebarBandInfo.fMask = RBBIM_STYLE | RBBIM_CHILDSIZE | RBBIM_HEADERSIZE;
    rebarBandInfo.fStyle &= ~RBBS_USECHEVRON;
    rebarBandInfo.cyMinChild = buttonRect.bottom+(HIWORD(padding)/2)+1;
    rebarBandInfo.cyMaxChild = buttonRect.bottom+(HIWORD(padding)/2)+1;
    rebarBandInfo.cxHeader = 6;
    SendMessage(rbWindow, RB_SETBANDINFO, (WPARAM) 0, (LPARAM) &rebarBandInfo);
}

void makeToolbarOverflow()
{
    HWND rbWindow, tbWindow;
    LONG_PTR style;
    REBARBANDINFO rebarBandInfo;
    DWORD size, padding;
    
    rbWindow = FindWindowEx(nppData._nppHandle, NULL, REBARCLASSNAME, NULL);
    tbWindow = FindWindowEx(rbWindow, NULL, TOOLBARCLASSNAME, NULL);
    
    style = GetWindowLongPtr(tbWindow, GWL_STYLE);
    style &= ~TBSTYLE_WRAPABLE;
    style |= TBSTYLE_TRANSPARENT;
    style |= CCS_ADJUSTABLE;
    SetWindowLongPtr(tbWindow, GWL_STYLE, style);
    
    size = (DWORD) SendMessage(tbWindow, TB_GETBUTTONSIZE, 0, 0);
    padding = (DWORD) SendMessage(tbWindow, TB_GETPADDING, 0, 0);
    
    rebarBandInfo.cbSize = g_rebarBandInfoSize;
    rebarBandInfo.fMask = RBBIM_STYLE | RBBIM_CHILDSIZE;
    SendMessage(rbWindow, RB_GETBANDINFO, (WPARAM) 0, (LPARAM) &rebarBandInfo);
    
    rebarBandInfo.cbSize = g_rebarBandInfoSize;
    rebarBandInfo.fMask = RBBIM_STYLE | RBBIM_CHILDSIZE | RBBIM_HEADERSIZE;
    rebarBandInfo.fStyle |= RBBS_USECHEVRON;
    rebarBandInfo.cyMinChild = HIWORD(size)+HIWORD(padding);
    rebarBandInfo.cyMaxChild = HIWORD(size)+HIWORD(padding);
    rebarBandInfo.cxHeader = 26;
    SendMessage(rbWindow, RB_SETBANDINFO, (WPARAM) 0, (LPARAM) &rebarBandInfo);
}

void adjustIdealSize()
{
    HWND rbWindow, tbWindow;
    SIZE tbMaxSize;
    REBARBANDINFO rebarBandInfo;
    
    rbWindow = FindWindowEx(nppData._nppHandle, NULL, REBARCLASSNAME, NULL);
    tbWindow = FindWindowEx(rbWindow, NULL, TOOLBARCLASSNAME, NULL);
    
    SendMessage(tbWindow, TB_GETMAXSIZE, (WPARAM) 0, (LPARAM) &tbMaxSize);
    
    rebarBandInfo.cbSize = g_rebarBandInfoSize;
    rebarBandInfo.fMask = RBBIM_IDEALSIZE;
    rebarBandInfo.cxIdeal = tbMaxSize.cx;
    
    SendMessage(rbWindow, RB_SETBANDINFO, (WPARAM) 0, (LPARAM) &rebarBandInfo);
}

void displayOverflowMenu(NMREBARCHEVRON *lpNmRebarChevron)
{
    HWND rbWindow, tbWindow;
    HMENU popupMenu;
    POINT popupPoint;
    RECT rbBandRect, tbButtonRect, intersectRect;
    TBBUTTON tbButton;
    TCHAR buffer[MAXSIZE];
    UINT menuStyle;
    int i, itemCount, buttonsOnToolbar;
    
    rbWindow = FindWindowEx(nppData._nppHandle, NULL, REBARCLASSNAME, NULL);
    tbWindow = FindWindowEx(rbWindow, NULL, TOOLBARCLASSNAME, NULL);
    
    // Create popup menu to show overflow buttons
    
    popupMenu = CreatePopupMenu();
    
    // Screen co-ordinates for Menu to be displayed - align with left of chevron (-12)
    
    popupPoint.x = lpNmRebarChevron->rc.left-11;
    popupPoint.y = lpNmRebarChevron->rc.bottom;
    ClientToScreen(tbWindow, &popupPoint);
    
    // Get band rectangle - subtract chevron width (-16)
    
    SendMessage(rbWindow, RB_GETRECT, (WPARAM) 0, (LPARAM) (LPPOINT) &rbBandRect);
    rbBandRect.right -= 16;
    
    // Add items to popup menu
    
    itemCount = 0;
    buttonsOnToolbar = (int) SendMessage(tbWindow, TB_BUTTONCOUNT, (WPARAM) 0, (LPARAM) 0);
    
    for (i = 0; i < buttonsOnToolbar; i++)
    {
        SendMessage(tbWindow, TB_GETBUTTON, (WPARAM) i, (LPARAM)(LPTBBUTTON) &tbButton);
        
        if (tbButton.fsStyle & BTNS_SEP)
        {
            // Add separator menu item
            
            if (itemCount > 0) AppendMenu(popupMenu, MF_SEPARATOR, 0, 0);
        }
        else
        {
            // Get button rectangle - offset by band gripper width (+12)
            
            SendMessage(tbWindow, TB_GETITEMRECT, (WPARAM) i, (LPARAM)(LPRECT) &tbButtonRect);
            tbButtonRect.left += 12;
            tbButtonRect.right += 12;
            
            // Check intersection of button and band rectangles
            
            IntersectRect(&intersectRect, &tbButtonRect, &rbBandRect);
            
            if (intersectRect.left != tbButtonRect.left || intersectRect.right != tbButtonRect.right ||
                intersectRect.top != tbButtonRect.top || intersectRect.bottom != tbButtonRect.bottom)
            {
                menuStyle = MF_STRING;
                menuStyle |= SendMessage(tbWindow, TB_ISBUTTONENABLED, (WPARAM) tbButton.idCommand, (LPARAM) 0) ? MF_ENABLED : MF_DISABLED|MF_GRAYED;
                menuStyle |= SendMessage(tbWindow, TB_ISBUTTONCHECKED, (WPARAM) tbButton.idCommand, (LPARAM) 0) ? MF_CHECKED : MF_UNCHECKED;
                
                SendMessage(tbWindow, TB_GETSTRING, (WPARAM) MAKEWPARAM(MAXSIZE,tbButton.iString), (LPARAM) buffer);
                
                AppendMenu(popupMenu, menuStyle, tbButton.idCommand, buffer);
                
                itemCount++;
            }
        }
    }
    
    // Display popup menu if at least one item has been added
    
    if (itemCount > 0) TrackPopupMenu(popupMenu, TPM_LEFTALIGN|TPM_TOPALIGN, popupPoint.x, popupPoint.y, 0, rbWindow, NULL);
    
    // Destroy popup menu
    
    DestroyMenu(popupMenu);
}

//
// Hash functions
//

DWORD calcButtonStringHash(TBBUTTON tbButton)
{
    HWND rbWindow, tbWindow;
    DWORD hash;
    TCHAR buffer[MAXSIZE];
    int i;
    
    rbWindow = FindWindowEx(nppData._nppHandle, NULL, REBARCLASSNAME, NULL);
    tbWindow = FindWindowEx(rbWindow, NULL, TOOLBARCLASSNAME, NULL);
    
    hash = 0;
    
    // Hash in button string
    
    SendMessage(tbWindow, TB_GETSTRING, (WPARAM) MAKEWPARAM(MAXSIZE,tbButton.iString), (LPARAM) buffer);
    for (i = 0; buffer[i] != 0; i++)
    {
        hash = ((hash << 5) - hash) + buffer[i]; /* hash * 31 + char */
    }
    
    hash |= HASHFLAG;  /* distinguish hash value from command identifier */
    
    return hash;
}

DWORD calcPluginButtonMenuHash(TBBUTTON tbButton)
{
    DWORD hash;
    TCHAR buffer[MAXSIZE];
    int i;
    
    hash = 0;
    
    // Hash in command menu string
    
    GetMenuString(g_hMainMenu, tbButton.idCommand, buffer, MAXSIZE, MF_BYCOMMAND);
    stripMenuString(buffer);
    for (i = 0; buffer[i] != 0; i++)
    {
        hash = ((hash << 5) - hash) + buffer[i]; /* hash * 31 + char */
    }
    
    // Hash in command parent menu string
    
    findPluginParentMenuString(g_hMainMenu, tbButton.idCommand, buffer, MAXSIZE);
    stripMenuString(buffer);
    for (i = 0; buffer[i] != 0; i++)
    {
        hash = ((hash << 5) - hash) + buffer[i]; /* hash * 31 + char */
    }
    
    hash |= HASHFLAG;  /* distinguish hash value from command identifier */
    
    return hash;
}

int findPluginParentMenuString(HMENU hMenu, UINT idCommand, LPTSTR lpString, int maxCount)
{
    HMENU hSubMenu;
    int i, count, result;
    
    count = GetMenuItemCount(hMenu);
    
    for (i = 0; i < count; i++)
    {
        hSubMenu = GetSubMenu(hMenu, i);
        if (hSubMenu == NULL)
        {
            if (GetMenuItemID(hMenu, i) == idCommand) return 1;  /* command identifier found at this level */
        }
        else
        {
            result = findPluginParentMenuString(hSubMenu, idCommand, lpString, maxCount);
            if (result == 1)  /* command identifier found one level below */
		    {
				GetMenuString(hMenu, i, lpString, maxCount, MF_BYPOSITION);
				return 0;  /* command identifier found */
			}
			if (result == 0) return result;  /* command identifier found */
        }
    }
    
    return -1;  /* command identifier not found */
}

//
// Custom button functions
//

int findCmdIDForMenuStrings(HMENU hMenu0, LPTSTR menuString0, LPTSTR menuString1, LPTSTR menuString2, LPTSTR menuString3)
{
    HMENU hMenu1, hMenu2, hMenu3, hMenu4;
    TCHAR buffer[MAXSIZE];
    int i, j, k, l, itemCount0, itemCount1, itemCount2, itemCount3;
    
    itemCount0 = GetMenuItemCount(hMenu0);
    for (i = 0; i < itemCount0; i++)
    {
        GetMenuString(hMenu0, i, buffer, MAXSIZE, MF_BYPOSITION);
        stripMenuString(buffer);
        if (lstrcmp(menuString0, buffer) == 0 && buffer[0] != 0)
        {
            hMenu1 = GetSubMenu(hMenu0, i);
            if (hMenu1 == NULL) return GetMenuItemID(hMenu0, i);
            
            itemCount1 = GetMenuItemCount(hMenu1);
            for (j = 0; j < itemCount1; j++)
            {
                GetMenuString(hMenu1, j, buffer, MAXSIZE, MF_BYPOSITION);
                stripMenuString(buffer);
                if (lstrcmp(menuString1, buffer) == 0 && buffer[0] != 0)
                {
                    hMenu2 = GetSubMenu(hMenu1, j);
                    if (hMenu2 == NULL) return GetMenuItemID(hMenu1, j);
                    
                    itemCount2 = GetMenuItemCount(hMenu2);
                    for (k = 0; k < itemCount2; k++)
                    {
                        GetMenuString(hMenu2, k, buffer, MAXSIZE, MF_BYPOSITION);
                        stripMenuString(buffer);
                        if (lstrcmp(menuString2, buffer) == 0 && buffer[0] != 0)
                        {
                            hMenu3 = GetSubMenu(hMenu2, k);
                            if (hMenu3 == NULL) return GetMenuItemID(hMenu2, k);
                            
                            itemCount3 = GetMenuItemCount(hMenu3);
                            for (l = 0; l < itemCount3; l++)
                            {
                                GetMenuString(hMenu3, l, buffer, MAXSIZE, MF_BYPOSITION);
                                stripMenuString(buffer);
                                if (lstrcmp(menuString3, buffer) == 0 && buffer[0] != 0)
                                {
                                    hMenu4 = GetSubMenu(hMenu3, l);
                                    if (hMenu4 == NULL) return GetMenuItemID(hMenu3, l);
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    
    return -1;
}

HBITMAP createBitmapForCustomButton(TCHAR text[])
{
    HDC hDC, hMemDC;
    HBITMAP hBitmap1, hBitmap2;
    HBRUSH hBrush;
    HFONT hFont;
    RECT rect;
    int label, count;
    unsigned int rh, rl, gh, gl, bh, bl;
    
    hBitmap1 = (HBITMAP) LoadImage((HINSTANCE) g_hModule, MAKEINTRESOURCE(IDB_CUSTOM_BACKGROUND16), IMAGE_BITMAP, 0, 0, (LR_DEFAULTSIZE | LR_LOADTRANSPARENT | LR_LOADMAP3DCOLORS));
    
    if (text[1] ==  0 || text[1] == (TCHAR) ':')  /* missing color code */
    {
        label = (text[1] == 0) ? 1 : 2;
        
        hBrush = CreateSolidBrush(RGB(128, 128, 128));
    }
    else if (text[2] ==  0 || text[2] == (TCHAR) ':')  /* letter color code */
    {
        label = (text[2] == 0) ? 2 : 3;
        
        if (text[1] == (TCHAR) 'S') hBrush = CreateSolidBrush(RGB(128, 128, 128));
        else if (text[1] == (TCHAR) 'R') hBrush = CreateSolidBrush(RGB(176, 48, 48));
        else if (text[1] == (TCHAR) 'G') hBrush = CreateSolidBrush(RGB(48, 144, 48));
        else if (text[1] == (TCHAR) 'B') hBrush = CreateSolidBrush(RGB(0, 80, 192));
        else if (text[1] == (TCHAR) 'C') hBrush = CreateSolidBrush(RGB(0, 160, 160));
        else if (text[1] == (TCHAR) 'M') hBrush = CreateSolidBrush(RGB(160, 64, 160));
        else if (text[1] == (TCHAR) 'Y') hBrush = CreateSolidBrush(RGB(176, 144, 0));
        else hBrush = CreateSolidBrush(RGB(128, 128, 128));
    }
    else  /* hex color code */
    {
        for (label = 1; text[label] != 0 && text[label] != (TCHAR) ':'; label++);
        label = (text[label] == (TCHAR) ':') ? label+1 : label;
        
        count = _stscanf_s(&text[1], TEXT("#%1x%1x%1x%1x%1x%1x"), &rh, &rl, &gh, &gl, &bh, &bl);
        hBrush = (count == 6) ? CreateSolidBrush(RGB(rh*16+rl, gh*16+gl, bh*16+bl)) : CreateSolidBrush(RGB(128, 128, 128));
    }
    
    text[label+2] = 0;  /* limit label to two characters */
    
    hFont = CreateFont(12, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, NONANTIALIASED_QUALITY, 0, TEXT("Lucida Console"));
    
    hDC = GetDC(NULL);  /* screen device context */
    hMemDC = CreateCompatibleDC(hDC);  /* Memory device context */
    
    SetBkMode(hMemDC, TRANSPARENT);
    SetBkColor(hMemDC, RGB(0, 0, 0));
    SetTextColor(hMemDC, RGB(255, 255, 255));
    
    SelectObject(hMemDC, hBitmap1);
    SelectObject(hMemDC, hFont);
    
    rect.left = 0;
    rect.right = 16;
    rect.top = 1;
    rect.bottom = 15;
    
    FillRect(hMemDC, &rect, hBrush);
    
    rect.left = (text[3] == 0) ? 1 : 0;
    rect.right = 16;
    rect.top = 3;
    rect.bottom = 16;
    
    DrawText(hMemDC, &text[label], -1, &rect, DT_SINGLELINE | DT_CENTER | DT_NOPREFIX);
    
    hBitmap2 = (HBITMAP) CopyImage(hBitmap1, IMAGE_BITMAP, 0, 0, 0);
    
    DeleteDC(hMemDC);
    
    DeleteObject(hBitmap1);
    DeleteObject(hFont);
    
    ReleaseDC(NULL, hDC);
    
    return hBitmap2;
}

HICON createIconForCustomButton(TCHAR text[])
{
    HDC hDC,hMemDC;
    HBITMAP hBitmap1, hBitmap2;
    HBRUSH hBrush;
    HFONT hFont;
    RECT rect;
    HIMAGELIST hImageList;
    HICON hIcon;
    int label, count;
    unsigned int rh, rl, gh, gl, bh, bl;
    
    hBitmap1 = (HBITMAP) LoadImage((HINSTANCE) g_hModule,MAKEINTRESOURCE(IDB_CUSTOM_BACKGROUND32), IMAGE_BITMAP, 0, 0, (LR_DEFAULTSIZE | LR_LOADTRANSPARENT | LR_LOADMAP3DCOLORS));
    
    if (text[1] ==  0 || text[1] == (TCHAR) ':')  /* missing color code */
    {
        label = (text[1] == 0) ? 1 : 2;
        
        hBrush = CreateSolidBrush(RGB(128, 128, 128));
    }
    else if (text[2] ==  0 || text[2] == (TCHAR) ':')  /* letter color code */
    {
        label = (text[2] == 0) ? 2 : 3;
        
        if (text[1] == (TCHAR) 'S') hBrush = CreateSolidBrush(RGB(128, 128, 128));
        else if (text[1] == (TCHAR) 'R') hBrush = CreateSolidBrush(RGB(176, 48, 48));
        else if (text[1] == (TCHAR) 'G') hBrush = CreateSolidBrush(RGB(48, 144, 48));
        else if (text[1] == (TCHAR) 'B') hBrush = CreateSolidBrush(RGB(0, 80, 192));
        else if (text[1] == (TCHAR) 'C') hBrush = CreateSolidBrush(RGB(0, 160, 160));
        else if (text[1] == (TCHAR) 'M') hBrush = CreateSolidBrush(RGB(160, 64, 160));
        else if (text[1] == (TCHAR) 'Y') hBrush = CreateSolidBrush(RGB(176, 144, 0));
        else hBrush = CreateSolidBrush(RGB(128, 128, 128));
    }
    else  /* hex color code */
    {
        for (label = 1; text[label] != 0 && text[label] != (TCHAR) ':'; label++);
        label = (text[label] == (TCHAR) ':') ? label+1 : label;
        
        count = _stscanf_s(&text[1], TEXT("#%1x%1x%1x%1x%1x%1x"), &rh, &rl, &gh, &gl, &bh, &bl);
        hBrush = (count == 6) ? CreateSolidBrush(RGB(rh*16+rl, gh*16+gl, bh*16+bl)) : CreateSolidBrush(RGB(128, 128, 128));
    }
    
    text[label+2] = 0;  /* limit label to two characters */
    
    hFont = CreateFont(24, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, PROOF_QUALITY, 0, TEXT("Lucida Console"));
    
    hDC = GetDC(NULL);  /* screen device context */
    hMemDC = CreateCompatibleDC(hDC);  /* Memory device context */
    
    SetBkMode(hMemDC,TRANSPARENT);
    SetBkColor(hMemDC, RGB(0, 0, 0));
    SetTextColor(hMemDC, RGB(254, 254, 254));
    
    SelectObject(hMemDC, hBitmap1);
    SelectObject(hMemDC, hFont);
    
    rect.left = 0;
    rect.right = 32;
    rect.top = 2;
    rect.bottom = 30;
    
    FillRect(hMemDC, &rect, hBrush);
    
    rect.left = (text[3] == 0) ? 1 : 0;
    rect.right = 32;
    rect.top = 5;
    rect.bottom = 32;
    
    DrawText(hMemDC, &text[label], -1, &rect, DT_SINGLELINE | DT_CENTER | DT_NOPREFIX);
    
    hBitmap2 = (HBITMAP) CopyImage(hBitmap1, IMAGE_BITMAP, 0, 0, 0);
    
    DeleteDC(hMemDC);
    
    DeleteObject(hBitmap1);
    DeleteObject(hFont);
    
    ReleaseDC(NULL, hDC);
    
    hImageList = ImageList_Create(32, 32, ILC_MASK | ILC_COLOR32, 1, 0);
    
    ImageList_AddMasked(hImageList, hBitmap2, CLR_DEFAULT);
    
    hIcon = ImageList_GetIcon(hImageList, 0, ILD_TRANSPARENT);
    
    return hIcon;
}

//
// Miscellaneous functions
//

void stripMenuString(LPTSTR lpString)
{
    int i, j;
    
    j = 0;
    for (i = 0; lpString[i] != 0; i++)
    {
        if (lpString[i] == '&');
        else if (lpString[i] == '\t') break;
        else lpString[j++] = lpString[i];
    }
    lpString[j] = 0;
}

int getCommCtrlMajorVersion()
{
    HINSTANCE hInstDLL;
    DLLGETVERSIONPROC pfnDllGetVersion;
    DLLVERSIONINFO versionInfo;
    
    hInstDLL = LoadLibrary(TEXT("comctl32.dll"));
    
    pfnDllGetVersion = (DLLGETVERSIONPROC) GetProcAddress(hInstDLL, "DllGetVersion");
    
    memset(&versionInfo, 0, sizeof(versionInfo));
    versionInfo.cbSize = sizeof(versionInfo);
    
    (*pfnDllGetVersion)(&versionInfo);
    
    FreeLibrary(hInstDLL);
    
    return (int) versionInfo.dwMajorVersion;
}

