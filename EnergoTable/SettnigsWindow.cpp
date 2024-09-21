#include "SettingsWindow.h"
#include "resource.h"

HINSTANCE parentInstance;
HWND mainWindow;
HWND window;
HWND hostLabel, portLabel, userLabel, passwordLabel, databaseLabel;
HWND hostInput, portInput, userInput, passwordInput, databaseInput;
HWND okButton, cancelButton;
WCHAR settingsWindowClass[100];
wchar_t address[17], port[5], user[10], password[10], database[20];
wchar_t filePath[255];
wchar_t file[200];
wchar_t section[9];
wchar_t parameter[10];

LRESULT CALLBACK settingsWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_CREATE:
        hostLabel = CreateWindowW(L"STATIC", L"Адрес", WS_CHILD | WS_VISIBLE | SS_RIGHT, 0, 11, 100, 20, hWnd, nullptr, parentInstance, nullptr);
        hostInput = CreateWindowW(L"EDIT", address, WS_TABSTOP | WS_CHILD | WS_VISIBLE | WS_BORDER, 140, 10, 100, 20, hWnd, nullptr, parentInstance, nullptr);
        portLabel = CreateWindowW(L"STATIC", L"Порт", WS_CHILD | WS_VISIBLE | SS_RIGHT, 0, 35, 100, 20, hWnd, nullptr, parentInstance, nullptr);
        portInput = CreateWindowW(L"EDIT", port, WS_TABSTOP | WS_CHILD | WS_VISIBLE | WS_BORDER, 140, 35, 100, 20, hWnd, nullptr, parentInstance, nullptr);
        userLabel = CreateWindowW(L"STATIC", L"Пользователь", WS_CHILD | WS_VISIBLE | SS_RIGHT, 0, 60, 100, 20, hWnd, nullptr, parentInstance, nullptr);
        userInput = CreateWindowW(L"EDIT", user, WS_TABSTOP | WS_CHILD | WS_VISIBLE | WS_BORDER, 140, 60, 100, 20, hWnd, nullptr, parentInstance, nullptr);
        passwordLabel = CreateWindowW(L"STATIC", L"Пароль", WS_CHILD | WS_VISIBLE | SS_RIGHT, 0, 90, 100, 20, hWnd, nullptr, parentInstance, nullptr);
        passwordInput = CreateWindowW(L"EDIT", password, WS_TABSTOP | WS_CHILD | WS_VISIBLE | WS_BORDER | ES_PASSWORD, 140, 90, 100, 20, hWnd, nullptr, parentInstance, nullptr);
        databaseLabel = CreateWindowW(L"STATIC", L"База данных", WS_CHILD | WS_VISIBLE | SS_RIGHT, 0, 120, 100, 20, hWnd, nullptr, parentInstance, nullptr);
        databaseInput = CreateWindowW(L"EDIT", database, WS_TABSTOP | WS_CHILD | WS_VISIBLE | WS_BORDER, 140, 120, 100, 20, hWnd, nullptr, parentInstance, nullptr);
        okButton = CreateWindowW(L"BUTTON", L"Сохранить", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, 150, 430, 100, 20, hWnd, (HMENU)IDB_OK_BUTTON, parentInstance, nullptr);
        cancelButton = CreateWindowW(L"BUTTON", L"Отмена", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, 270, 430, 100, 20, hWnd, (HMENU)IDB_CANCLEL_BUTTON, parentInstance, nullptr);
        break;
    case WM_COMMAND:
    {
        int wmId = LOWORD(wParam);
        switch (wmId) {
        case IDB_OK_BUTTON:
            save();
            break;
        case IDB_CANCLEL_BUTTON:
            cancel();
            break;
        default:
            return DefWindowProc(hWnd, message, wParam, lParam);
        }
        break;
    }
    case WM_DESTROY:
        EnableWindow(mainWindow, true);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}
void initSettingsWindow(HINSTANCE hInstance, HWND parent) {

        WNDCLASSEXW wcex;

        wsprintf(settingsWindowClass, L"%s", L"SettingsWindow");

        wcex.cbSize = sizeof(WNDCLASSEX);

        wcex.style = CS_HREDRAW | CS_VREDRAW;
        wcex.lpfnWndProc = settingsWndProc;
        wcex.cbClsExtra = 0;
        wcex.cbWndExtra = 0;
        wcex.hInstance = hInstance;
        wcex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ENERGOTABLE));
        wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
        wcex.hbrBackground = CreateSolidBrush(RGB(240, 240, 240));
        wcex.lpszMenuName = NULL;
        wcex.lpszClassName = settingsWindowClass;
        wcex.hIconSm = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

        RegisterClassExW(&wcex);

        parentInstance = hInstance;
        mainWindow = parent;
        window = NULL;

        GetCurrentDirectory(255, filePath);
        wsprintf(file, L"%s%s", filePath, L"\\settings.ini");
        wsprintf(section, L"%s", L"Settings");
        GetPrivateProfileString(section, L"host", L"127.0.0.1", address, 17, file);
        GetPrivateProfileString(section, L"port", L"5432", port, 5, file);
        GetPrivateProfileString(section, L"user", L"postgres", user, 10, file);
        GetPrivateProfileString(section, L"password", L"12345678", password, 10, file);
        GetPrivateProfileString(section, L"database", L"EnergoCenter", database, 20, file);
}
void showSettingsWindow() {
    window = CreateWindow(settingsWindowClass, L"Settings", WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU, 0, 0, 400, 500, mainWindow, NULL, parentInstance, NULL);
    ShowWindow(window, SW_SHOW);
    UpdateWindow(window);
    EnableWindow(mainWindow, false);
}
void save() {
    GetWindowText(hostInput, address, 17);
    GetWindowText(portInput, port, 5);
    GetWindowText(userInput, user, 10);
    GetWindowText(passwordInput, password, 10);
    GetWindowText(databaseInput, database, 20);
    wsprintf(parameter, L"%s", L"host");
    WritePrivateProfileString(section, parameter, address, file);
    wsprintf(parameter, L"%s", L"port");
    WritePrivateProfileString(section, parameter, port, file);
    wsprintf(parameter, L"%s", L"user");
    WritePrivateProfileString(section, parameter, user, file);
    wsprintf(parameter, L"%s", L"password");
    WritePrivateProfileString(section, parameter, password, file);
    wsprintf(parameter, L"%s", L"database");
    WritePrivateProfileString(section, parameter, database, file);
    DestroyWindow(window);
    BringWindowToTop(mainWindow);
}
void cancel() {
    DestroyWindow(window);
    BringWindowToTop(mainWindow);
}
wchar_t* getHost() {
    return address;
}
wchar_t* getPort() {
    return port;
}
wchar_t* getUser() {
    return user;
}
wchar_t* getPassword() {
    return password;
}
wchar_t* getDatabaseName() {
    return database;
}