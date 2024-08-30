// This file is part of notepad++
// Copyright (C) 2003 Don HO <donho@altern.org>

// Modifications to this file for Customize Toolbar plugin
// Copyright (C) 2011-2021 DW-dev (dw-dev@gmx.com)
// Last Edit - 26 Jul 2021

#ifndef PLUGINDEFINITION_H
#define PLUGINDEFINITION_H

//
// All difinitions of plugin interface
//
#include "PluginInterface.h"

//-------------------------------------//
//-- STEP 1. DEFINE YOUR PLUGIN NAME --//
//-------------------------------------//
// Here define your plugin name
//
const TCHAR NPP_PLUGIN_NAME[] = TEXT("Customize Toolbar");

//-----------------------------------------------//
//-- STEP 2. DEFINE YOUR PLUGIN COMMAND NUMBER --//
//-----------------------------------------------//
//
// Here define the number of your plugin commands
//
const int nbFunc = 9;


//
// Initialization of your plugin data
// It will be called while plugin loading
//
void pluginInit(HANDLE hModule);

//
// Cleaning of your plugin
// It will be called while plugin unloading
//
void pluginCleanUp();

//
//Initialization of your plugin commands
//
void commandMenuInit();

//
//Clean up your plugin commands allocation (if any)
//
void commandMenuCleanUp();

//
// Function which sets your command
//
bool setCommand(size_t index, TCHAR *cmdName, PFUNCPLUGINCMD pFunc, ShortcutKey *sk = NULL, bool check0nInit = false);


//
// Your plugin command functions
//
void addMenuCommands();
void addToolbarButtons();
void afterNppReady();
void bufferActivated();
void beforeNppShutdown();
void customizeToolbar();
void customButtons();
void wrapToolbar();
void helpOverview();
void helpCustomButtons();
void resourceUsage();

#endif //PLUGINDEFINITION_H
