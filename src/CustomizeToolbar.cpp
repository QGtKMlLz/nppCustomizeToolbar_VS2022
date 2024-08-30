// This file is part of notepad++
// Copyright (C) 2003 Don HO <donho@altern.org>

// Modifications to this file for Customize Toolbar plugin
// Copyright (C) 2011-2021 DW-dev (dw-dev@gmx.com)
// Last Edit - 02 Jul 2021

#include "PluginDefinition.h"

extern FuncItem funcItem[nbFunc];
extern NppData nppData;

/* global values */

HANDLE		g_hModule;

// Sequence of events is:
//
//      DllMain()       DLL_PROCESS_ATTACH      >>  pluginInit()
//      setInfo()                               >>  commandMenuInit()       addMenuCommands()
//      beNotified()    NPPN_TBMODIFICATION     >>                          addToolbarButtons()
//      beNotified()    NPPN_READY              >>                          afterNppReady()
//      beNotified()    NPPN_BUFFERACTIVATED    >>                          bufferActivated()
//      beNotified()    NPPN_SHUTDOWN           >>  commandMenuCleanUp()    beforeNppShutdown()
//      DllMain()       DLL_PROCESS_DETACH      >>  pluginCleanUp()

/* functions */

BOOL APIENTRY DllMain( HANDLE hModule,
                       DWORD  reasonForCall,
                       LPVOID lpReserved )
{
    g_hModule = hModule;

    switch (reasonForCall)
    {
        case DLL_PROCESS_ATTACH:
            pluginInit(hModule);
            break;

        case DLL_PROCESS_DETACH:
            pluginCleanUp();
            break;

        case DLL_THREAD_ATTACH:
            break;

        case DLL_THREAD_DETACH:
            break;
    }

    return TRUE;
}

extern "C" __declspec(dllexport) void setInfo(NppData notpadPlusData)
{
    nppData = notpadPlusData;
    commandMenuInit();
    addMenuCommands();
}

extern "C" __declspec(dllexport) const TCHAR * getName()
{
    return NPP_PLUGIN_NAME;
}

extern "C" __declspec(dllexport) FuncItem * getFuncsArray(int *nbF)
{
    *nbF = nbFunc;
    return funcItem;
}

extern "C" __declspec(dllexport) void beNotified(SCNotification *notifyCode)
{
    if (notifyCode->nmhdr.hwndFrom == nppData._nppHandle)
    {
        switch (notifyCode->nmhdr.code)
        {
            case NPPN_TBMODIFICATION:
            {
                addToolbarButtons();
            }
            break;

            case NPPN_READY:
            {
                afterNppReady();
            }
            break;

            case NPPN_SHUTDOWN:
            {
                commandMenuCleanUp();
                beforeNppShutdown();
            }
            break;

            default:
            return;
        }
    }
}

// Here you can process the Npp Messages
// I will make the messages accessible little by little, according to the need of plugin development.
// Please let me know if you need to access to some messages :
// http://sourceforge.net/forum/forum.php?forum_id=482781
//
extern "C" __declspec(dllexport) LRESULT messageProc(UINT Message, WPARAM wParam, LPARAM lParam)
{
    /*
    if (Message == WM_MOVE)
    {
	    ::MessageBox(NULL, "move", "", MB_OK);
    }
    */
    return TRUE;
}

#ifdef UNICODE
extern "C" __declspec(dllexport) BOOL isUnicode()
{
    return TRUE;
}
#endif //UNICODE

