#pragma once
#include "framework.h"
void initSettingsWindow(HINSTANCE hInstance, HWND parent);
void showSettingsWindow();
void save();
void cancel();
wchar_t* getHost();
wchar_t* getPort();
wchar_t* getUser();
wchar_t* getPassword();
wchar_t* getDatabaseName();