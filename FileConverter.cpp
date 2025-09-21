// FileConverter.cpp : Defines the entry point for the application.
//

#include "framework.h"
#include "FileConverter.h"
#include "shobjidl_core.h"
#include <iostream>
#include <winuser.h>
#include <windows.h>
#include <string>
#include <vector>
#include <cstdlib>
#include <sstream>
#include <locale>
#include <codecvt>
#include <Urlmon.h>

#define MAX_LOADSTRING 100
#define CONVERT_BUTTON 1223
#define FILE_BUTTON 4407

HINSTANCE hInst;                                // current instance
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name
HWND hWnd;                                      // the main window
std::vector<std::wstring> path;                 // the file path
    
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

    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_FILECONVERTER, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    if (!InitInstance (hInstance, nCmdShow))
    {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_FILECONVERTER));

    MSG msg;

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


/**
 * @brief Open a dialog to select item(s) or folder(s).
 * @param paths Specifies the reference to the string vector that will receive the file or folder path(s). [IN]
 * @param selectFolder Specifies whether to select folder(s) rather than file(s). (optional)
 * @param multiSelect Specifies whether to allow the user to select multiple items. (optional)
 * @note If no item(s) were selected, the function still returns true, and the given vector is unmodified.
 * @note `<windows.h>`, `<string>`, `<vector>`, `<shobjidl.h>`
 * @return Returns true if all the operations are successfully performed, false otherwise.
 */
bool OpenFileDialog(std::vector<std::wstring>& paths, bool selectFolder, bool multiSelect)
{
    HRESULT hrInit = CoInitialize(NULL); // Initialize COM
    IFileOpenDialog* p_file_open = nullptr;
    bool are_all_operation_success = false;
    while (!are_all_operation_success)
    {
        HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_ALL,
            IID_IFileOpenDialog, reinterpret_cast<void**>(&p_file_open));
        if (FAILED(hr))
            break;

        if (selectFolder || multiSelect)
        {
            FILEOPENDIALOGOPTIONS options = 0;
            hr = p_file_open->GetOptions(&options);
            if (FAILED(hr))
                break;

            if (selectFolder)
                options |= FOS_PICKFOLDERS;
            if (multiSelect)
                options |= FOS_ALLOWMULTISELECT;

            hr = p_file_open->SetOptions(options);
            if (FAILED(hr))
                break;
        }

        hr = p_file_open->Show(NULL);

        if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED)) 
        {
            are_all_operation_success = true;
            break;
        }
        else if (FAILED(hr))
            break;

        IShellItemArray* p_items;
        hr = p_file_open->GetResults(&p_items);
        if (FAILED(hr))
            break;
        DWORD total_items = 0;
        hr = p_items->GetCount(&total_items);
        if (FAILED(hr))
            break;

        for (int i = 0; i < static_cast<int>(total_items); ++i)
        {
            IShellItem* p_item;
            p_items->GetItemAt(i, &p_item);
            if (SUCCEEDED(hr))
            {
                PWSTR path;
                hr = p_item->GetDisplayName(SIGDN_FILESYSPATH, &path);
                if (SUCCEEDED(hr))
                {
                    paths.push_back(path);
                    CoTaskMemFree(path);
                }
                p_item->Release();
            }
        }

        p_items->Release();
        are_all_operation_success = true;
    }

    if (p_file_open)
        p_file_open->Release();
    if (SUCCEEDED(hrInit)) CoUninitialize(); // Uninitialize COM
    return are_all_operation_success;
}


std::wstring StringToWString(const std::string& s) {
    int len;
    int slength = (int)s.length() + 1;
    len = MultiByteToWideChar(CP_UTF8, 0, s.c_str(), slength, 0, 0);
    std::wstring ws(len, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, s.c_str(), slength, &ws[0], len);
    // Remove the trailing null character
    if (!ws.empty() && ws.back() == L'\0') ws.pop_back();
    return ws;
}

char* wstringVectorToChar(const std::vector<std::wstring>& wvec) {
    // Convert each wstring to UTF-8 string
    std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
    std::ostringstream oss;

    for (const auto& wstr : wvec) {
        oss << converter.to_bytes(wstr);
    }

    std::string combined = oss.str();
    char* result = new char[combined.size() + 1];
    std::copy(combined.begin(), combined.end(), result);
    result[combined.size()] = '\0'; 

    return result;
}


/**
 * @brief Open a dialog to save an item.
 * @param path Specifies the reference to the string that will receive the target save path. [IN]
 * @param defaultFileName Specifies the default save file name. (optional)
 * @param pFilterInfo Specifies the pointer to the pair that contains filter information. (optional)
 * @note If no path was selected, the function still returns true, and the given string is unmodified.
 * @note `<windows.h>`, `<string>`, `<vector>`, `<shobjidl.h>`
 * @return Returns true if all the operations are successfully performed, false otherwise.
 */
bool SaveFileDialog(std::wstring& path, std::wstring defaultFileName = L"", std::pair<COMDLG_FILTERSPEC*, int>* pFilterInfo = nullptr)
{
    IFileSaveDialog* p_file_save = nullptr;
    bool are_all_operation_success = false;
    while (!are_all_operation_success)
    {
        HRESULT hr = CoCreateInstance(CLSID_FileSaveDialog, NULL, CLSCTX_ALL,
            IID_IFileSaveDialog, reinterpret_cast<void**>(&p_file_save));
        if (FAILED(hr))
            break;

        if (!pFilterInfo)
        {
            COMDLG_FILTERSPEC save_filter[1];
            save_filter[0].pszName = L"All files";
            save_filter[0].pszSpec = L"*.*";
            hr = p_file_save->SetFileTypes(1, save_filter);
            if (FAILED(hr))
                break;
            hr = p_file_save->SetFileTypeIndex(1);
            if (FAILED(hr))
                break;
        }
        else
        {
            hr = p_file_save->SetFileTypes(pFilterInfo->second, pFilterInfo->first);
            if (FAILED(hr))
                break;
            hr = p_file_save->SetFileTypeIndex(1);
            if (FAILED(hr))
                break;
        }

        if (!defaultFileName.empty())
        {
            hr = p_file_save->SetFileName(defaultFileName.c_str());
            if (FAILED(hr))
                break;
        }

        hr = p_file_save->Show(NULL);
        if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED)) // No item was selected.
        {
            are_all_operation_success = true;
            break;
        }
        else if (FAILED(hr))
            break;

        IShellItem* p_item;
        hr = p_file_save->GetResult(&p_item);
        if (FAILED(hr))
            break;

        PWSTR item_path;
        hr = p_item->GetDisplayName(SIGDN_FILESYSPATH, &item_path);
        if (FAILED(hr))
            break;
        path = item_path;
        CoTaskMemFree(item_path);
        p_item->Release();

        are_all_operation_success = true;
    }

    if (p_file_save)
        p_file_save->Release();
    return are_all_operation_success;
}


//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE: Processes messages for the main window.
//
//  WM_COMMAND  - process the application menu
//  WM_PAINT    - Paint the main window
//  WM_DESTROY  - post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_COMMAND:
    {
        if (LOWORD(wParam) == CONVERT_BUTTON && HIWORD(wParam) == BN_CLICKED) {

            MessageBox(hWnd, L"Convert", L"Hello", NULL);

            std::string cmd = "python fileConverter.py ";
            char* pathStr = wstringVectorToChar(path);
            cmd += pathStr;
            delete[] pathStr;

            FILE* pipe = _popen(cmd.c_str(), "r");
            if (!pipe) return 1;

            char buffer[128];
            std::string result;
            while (fgets(buffer, sizeof(buffer), pipe) != nullptr) {
                result += buffer;
            }
            _pclose(pipe);



            // After reading the result from the Python script
            result.erase(0, result.find_first_not_of(" \n\r\t")); // Trim leading whitespace
            result.erase(result.find_last_not_of(" \n\r\t") + 1); // Trim trailing whitespace

            if (result.rfind("http://", 0) == 0 || result.rfind("https://", 0) == 0) {
                // Valid URL, proceed with download
                std::wstring wurl = StringToWString(result);
                LPCWSTR url = wurl.c_str();
                std::wstring savePath;
                SaveFileDialog(savePath, L"convertedfile");
                HRESULT hr = URLDownloadToFile(NULL, url, savePath.c_str(), 0, NULL);
                // Handle hr as needed
            }
            else {
                MessageBox(hWnd, L"Python did not return a valid URL.", L"Error", MB_ICONERROR);
            }


            std::wstring wurl = StringToWString(result);
            LPCWSTR url = wurl.c_str();

            std::wstring savePath;
            SaveFileDialog(savePath, L"convertedfile");

            HRESULT hr = URLDownloadToFile(NULL, url, savePath.c_str(), 0, NULL);

            if (SUCCEEDED(hr)) {
                MessageBox(hWnd, L"File downloaded successfully.", L"Success", NULL);

                //std::cout << "File downloaded successfully." << std::endl;
            }
            else {
                MessageBox(hWnd, L"Error downloading file", L"ERROR", NULL);
               // std::cerr << "Error downloading file. HRESULT: " << hr << std::endl;
            }

            /*  Test code for saving
            FILE* pipe = _popen("python test.py ", "r");
            if (!pipe) return 1;

            char buffer[128];
            std::string result;
            while (fgets(buffer, sizeof(buffer), pipe) != nullptr) {
                result += buffer;
            }
            _pclose(pipe);

            std::cout << "Python returned: " << result << std::endl;

            HRESULT hr;
            std::wstring wurl = StringToWString(result);
            LPCWSTR url = wurl.c_str();
            // LPCWSTR localPath = L"converted-file";
            std::wstring savePath;
            SaveFileDialog(savePath, L"convertedfile");

            hr = URLDownloadToFile(NULL, url, savePath.c_str(), 0, NULL);

            if (SUCCEEDED(hr)) {
                std::cout << "File downloaded successfully." << std::endl;
            }
            else {
                std::cerr << "Error downloading file. HRESULT: " << hr << std::endl;
              }
              */

            /*
            // old version
            cmd += pathStr;
            system(cmd.c_str());
            std::string echo = "echo ";
            echo += pathStr;
            echo += "\npause";
            system(echo.c_str());
            delete[] pathStr;
            */


        }

        if (LOWORD(wParam) == FILE_BUTTON && HIWORD(wParam) == BN_CLICKED) {
            MessageBox(hWnd, L"Open", L"Hello", NULL);

            OpenFileDialog(path, false, false);

        }


        int wmId = LOWORD(wParam);

        switch (wmId)
        {
        case IDM_ABOUT:
            DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
            break;
        case IDM_EXIT:
            DestroyWindow(hWnd);
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
        // TODO: Add any drawing code that uses hdc here...
        EndPaint(hWnd, &ps);
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

//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
    wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_FILECONVERTER));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_FILECONVERTER);
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}


//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   hInst = hInstance; // Store instance handle in our global variable

   hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);


   HWND convertButton = CreateWindow(
       L"BUTTON",       // Predefined class; Unicode assumed 
       L"CONVERT",      // Button text 
       WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,  // Styles 
       120,         // x position 
       10,          // y position 
       100,         // Button width
       100,         // Button height
       hWnd,        // Parent window
       (HMENU) CONVERT_BUTTON,  // No menu.
       (HINSTANCE)GetWindowLongPtr(hWnd, GWLP_HINSTANCE),
       NULL         // Pointer not needed.
   );

   HWND chooseButton = CreateWindow(
       L"BUTTON",       // Predefined class; Unicode assumed 
       L"CHOOSE FILE",      // Button text 
       WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,  // Styles 
       10,         // x position 
       10,          // y position 
       100,         // Button width
       100,         // Button height
       hWnd,        // Parent window
       (HMENU)FILE_BUTTON,  // No menu.
       (HINSTANCE)GetWindowLongPtr(hWnd, GWLP_HINSTANCE),
       NULL         // Pointer not needed.
   );


   if (!hWnd)
   {
      return FALSE;
   }

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   return TRUE;
}

// Message handler for about box.
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
