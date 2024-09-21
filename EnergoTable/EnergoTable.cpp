#include "framework.h"
#include "EnergoTable.h"
#include "SettingsWindow.h"

#define MAX_LOADSTRING 100

HINSTANCE hInst;
WCHAR szTitle[MAX_LOADSTRING];
WCHAR szWindowClass[MAX_LOADSTRING];
INITCOMMONCONTROLSEX icex;
HWND dateSelector;
HWND startButton;
SYSTEMTIME timeSelect;
wchar_t month[3];
wchar_t year[5];
STARTUPINFO si;
PROCESS_INFORMATION pi;
wchar_t command[255];
wchar_t path[255];

ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // TODO: Разместите код здесь.

    // Инициализация глобальных строк
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_ENERGOTABLE, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    // Выполнить инициализацию приложения:
    if (!InitInstance (hInstance, nCmdShow))
    {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_ENERGOTABLE));

    MSG msg;

    // Цикл основного сообщения:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    return (int) msg.wParam;
}

ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
    wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ENERGOTABLE));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_ENERGOTABLE);
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   hInst = hInstance;

   HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, CW_USEDEFAULT, 500, 100, nullptr, nullptr, hInstance, nullptr);

   if (!hWnd)
   {
      return FALSE;
   }

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   return TRUE;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    LPITEMIDLIST pidl;
    BROWSEINFO bi;
    switch (message)
    {
    case WM_CREATE:
        initSettingsWindow(hInst, hWnd);
        icex.dwSize = sizeof(icex);
        icex.dwICC = ICC_DATE_CLASSES;
        InitCommonControlsEx(&icex);
        dateSelector = CreateWindowEx(0, DATETIMEPICK_CLASS, TEXT("Select date"), WS_BORDER | WS_CHILD | WS_VISIBLE | DTS_SHOWNONE,
            20, 5, 150, 25, hWnd, nullptr, hInst, nullptr);
        startButton = CreateWindow(L"BUTTON", L"Начать", WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
            180, 5, 100, 25, hWnd, (HMENU)IDB_START_BUTTON, hInst, nullptr);
        GetCurrentDirectory(255, path);
        break;
    case WM_COMMAND:
    {
        int wmId = LOWORD(wParam);
        switch (wmId)
        {
        case IDM_ABOUT:
            DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
            break;
        case IDM_EXIT:
            DestroyWindow(hWnd);
            break;
        case IDB_START_BUTTON:
            TCHAR filePath[MAX_PATH];
            bi = { 0 };
            bi.lpszTitle = L"Browse for folder...";
            bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;
            //bi.lpfn = BrowseCallbackProc;
            bi.lParam = (LPARAM)path;

            pidl = SHBrowseForFolder(&bi);

            if (pidl != 0)
            {
                SHGetPathFromIDList(pidl, filePath);

                IMalloc* imalloc = 0;
                if (SUCCEEDED(SHGetMalloc(&imalloc)))
                {
                    imalloc->Free(pidl);
                    imalloc->Release();
                }
                SendMessage(dateSelector, DTM_GETSYSTEMTIME, NULL, (LPARAM)&timeSelect);
                wsprintf(month, L"%hu", timeSelect.wMonth);
                wsprintf(year, L"%hu", timeSelect.wYear);
                ZeroMemory(&si, sizeof(si));
                si.cb = sizeof(si);
                ZeroMemory(&pi, sizeof(pi));
                wsprintf(command, L"%s\\Python\\python.exe %s\\pgtoexcel.py -h %s -p %s -u %s -pw %s -db %s -f %s -m %s -y %s",\
                    path, path, getHost(), getPort(), getUser(), getPassword(), getDatabaseName(), filePath, month, year);
                if (!CreateProcess(NULL, command,        // Command line
                    NULL,           // Process handle not inheritable
                    NULL,           // Thread handle not inheritable
                    FALSE,          // Set handle inheritance to FALSE
                    0,              // No creation flags
                    NULL,           // Use parent's environment block
                    NULL,           // Use parent's starting directory 
                    &si,            // Pointer to STARTUPINFO structure
                    &pi)           // Pointer to PROCESS_INFORMATION structure
                    )
                    {
                        SetWindowText(hWnd, L"Error create process");
                    }
                WaitForSingleObject(pi.hProcess, INFINITE);
                CloseHandle(pi.hProcess);
                CloseHandle(pi.hThread);
            }
            break;
        case IDM_SETTINGS:
            showSettingsWindow();
            break;
        default:
            return DefWindowProc(hWnd, message, wParam, lParam);
        }
    }
    break;
    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);
        EndPaint(hWnd, &ps);
    }
    break;
    case WM_NOTIFY:
        switch (((LPNMHDR)lParam)->code) {
        case DTN_DATETIMECHANGE:
            SendMessage(dateSelector, DTM_GETSYSTEMTIME, NULL, (LPARAM)&timeSelect);
            break;
        }
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}
