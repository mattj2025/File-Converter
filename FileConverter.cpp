/**
* File Converter Win32 Program
* 
* https://github.com/mattj2025/File-Converter
* 
*/

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
#include <commctrl.h> 

#define MAX_LOADSTRING 100
#define CONVERT_BUTTON 1223
#define FILE_BUTTON 4407
#define FILE_INPUT 2016
#define TYPE_INPUT 80085

HINSTANCE hInst;                                // current instance
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name
HWND hWnd;                                      // the main window
std::vector<std::wstring> path;                 // the file path
std::string outputFileExtension = "Other";      // the file type for output
HWND fileLocation;                              // file location input obj
HWND hEdit;                                     // type export input obj

const LPCWSTR LfileExtensions[13] = {
    TEXT("JPG"), TEXT("PDF"), TEXT("MP3"), TEXT("HEIC"),
    TEXT("Webp"), TEXT("WORD"), TEXT("MP4"), TEXT("HTML"),
    TEXT("GIF"), TEXT("JPEG"), TEXT("DOCX"), TEXT("PNG"),
    TEXT("Other")
};
const std::string SfileExtensions[13] = {
     "JPG",  "PDF",  "MP3",  "HEIC",
     "Webp",  "WORD",  "MP4",  "HTML",
     "GIF",  "JPEG",  "DOCX",  "PNG",
     "Other"
};
    
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
    HRESULT hrInit = CoInitialize(NULL);
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

    SetWindowText(fileLocation, StringToWString(wstringVectorToChar(path)).c_str());

    if (p_file_open)
        p_file_open->Release();
    if (SUCCEEDED(hrInit)) CoUninitialize(); // Uninitialize COM
    return are_all_operation_success;
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
            // TODO: Change save_filter
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

wchar_t* stringToWchar_t(const std::string& s) {
    // Determine the required buffer size
    size_t requiredSize;
    mbstowcs_s(&requiredSize, nullptr, 0, s.c_str(), _TRUNCATE);

    // Allocate memory for the wide character string
    // Add 1 for the null terminator
    wchar_t* wcString = new wchar_t[requiredSize];

    // Perform the conversion
    mbstowcs_s(nullptr, wcString, requiredSize, s.c_str(), _TRUNCATE);

    return wcString;
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

            // TODO: update outputFileExtension

            std::string cmd = "python fileConverter.py \"";
 
            int len = GetWindowTextLength(fileLocation);
            wchar_t* pathStr = new wchar_t[len + 1];

            GetWindowText(fileLocation, pathStr, len + 1);

            char narrow_buffer[256]; 
            size_t converted_chars = 0; 

            size_t required_size;
            wcstombs_s(&required_size, NULL, 0, pathStr, _TRUNCATE);

            errno_t err = wcstombs_s(
                &converted_chars,     
                narrow_buffer,        
                sizeof(narrow_buffer), 
                pathStr,              
                _TRUNCATE             
            );

            cmd += narrow_buffer;
            cmd += "\" ";

            if (outputFileExtension != "Other") {
                cmd += outputFileExtension;
            }
            else {
                int len = GetWindowTextLength(hEdit);
                wchar_t* type = new wchar_t[len + 1];
                GetWindowText(hEdit, type, len);

                char narrow_buffer[256];
                size_t converted_chars = 0;

                size_t required_size;
                wcstombs_s(&required_size, NULL, 0, pathStr, _TRUNCATE);

                errno_t err = wcstombs_s(
                    &converted_chars,
                    narrow_buffer,
                    sizeof(narrow_buffer),
                    pathStr,
                    _TRUNCATE
                );

                cmd += narrow_buffer;
            }
            delete[] pathStr;
               
            std::wstring myWString(cmd.begin(), cmd.end());
            LPCWSTR lpcwstr = myWString.c_str();
            MessageBox(hWnd, lpcwstr, NULL, NULL);

            FILE* pipe = _popen(cmd.c_str(), "r");
            if (!pipe) return 1;

            char buffer[128];
            std::string result;
            while (fgets(buffer, sizeof(buffer), pipe) != nullptr) {
                result += buffer;
            }
            _pclose(pipe);


            // trim whitespace
            result.erase(0, result.find_first_not_of(" \n\r\t")); 
            result.erase(result.find_last_not_of(" \n\r\t") + 1);

            if (result.rfind("http://", 0) == 0 || result.rfind("https://", 0) == 0) {

                std::wstring wurl = StringToWString(result);
                LPCWSTR url = wurl.c_str();
                std::wstring savePath;
                SaveFileDialog(savePath, L"convertedfile");
                HRESULT hr = URLDownloadToFile(NULL, url, savePath.c_str(), 0, NULL);

            }
            else {
                MessageBox(hWnd, stringToWchar_t(result), L"Python did not return a valid URL", MB_ICONERROR);
            }


            std::wstring wurl = StringToWString(result);
            LPCWSTR url = wurl.c_str();

            std::wstring savePath;
            SaveFileDialog(savePath, L"convertedfile");

            HRESULT hr = URLDownloadToFile(NULL, url, savePath.c_str(), 0, NULL);

            if (SUCCEEDED(hr)) {
                MessageBox(hWnd, L"File downloaded successfully.", L"Success", NULL);
            }
            else {
                MessageBox(hWnd, L"Error downloading file", L"ERROR", NULL);
            }


        }

        if (HIWORD(wParam) == CBN_SELCHANGE) {
            int ItemIndex = SendMessage((HWND)lParam, (UINT)CB_GETCURSEL, (WPARAM)0, (LPARAM)0);
            outputFileExtension = SfileExtensions[ItemIndex];
        }


        if (LOWORD(wParam) == FILE_BUTTON && HIWORD(wParam) == BN_CLICKED) {

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
        // No custom drawing needed for "Output Type:" label here.
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

// Set Segoe UI font for the main window and all child controls
void SetModernFont(HWND hwnd) {
    HFONT hFont = CreateFontW(
        18, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI"
    );
    SendMessage(hwnd, WM_SETFONT, (WPARAM)hFont, TRUE);

    // Enumerate all child windows and set their font
    HWND child = GetWindow(hwnd, GW_CHILD);
    while (child) {
        SendMessage(child, WM_SETFONT, (WPARAM)hFont, TRUE);
        child = GetWindow(child, GW_HWNDNEXT);
    }
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

    int winWidth = 480;
    int winHeight = 260;

    hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, 0, winWidth, winHeight, nullptr, nullptr, hInstance, nullptr);

    int padding = 20;
    int controlHeight = 32;
    int labelHeight = 22;
    int buttonHeight = 38;
    int buttonWidth = 140;
    int inputWidth = 260;
    int comboWidth = 160;

    // File Location 
    fileLocation = CreateWindowEx(
        WS_EX_CLIENTEDGE,
        TEXT("EDIT"),
        TEXT("File Location"),
        WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
        padding, padding + labelHeight,     // x, y
        inputWidth, controlHeight,          // width, height
        hWnd,
        (HMENU)FILE_INPUT,
        GetModuleHandle(NULL),
        NULL
    );

    // CHOOSE FILE button (left side, below file location)
    HWND chooseButton = CreateWindowEx(
        0,
        L"BUTTON",
        L"Browse...",
        WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
        padding, padding + labelHeight + controlHeight, // x, y
        buttonWidth, buttonHeight,     // width, height
        hWnd,
        (HMENU)FILE_BUTTON,
        (HINSTANCE)GetWindowLongPtr(hWnd, GWLP_HINSTANCE),
        NULL
    );

    // Output Type label (above ComboBox)
    HWND typeLabel = CreateWindowEx(
        0,
        L"STATIC",
        L"Output Type:",
        WS_CHILD | WS_VISIBLE,
        padding + inputWidth + 10, padding,     // x, y
        comboWidth, labelHeight,     // width, height
        hWnd,
        NULL,
        (HINSTANCE)GetWindowLongPtr(hWnd, GWLP_HINSTANCE),
        NULL
    );

    // ComboBox (right side, below label)
    HWND typeSelector = CreateWindowEx(
        0,
        WC_COMBOBOX,
        L"",
        CBS_DROPDOWN | CBS_HASSTRINGS | WS_CHILD | WS_OVERLAPPED | WS_VISIBLE,
        padding + inputWidth + 10, padding + labelHeight,     // x, y
        comboWidth, 180,    // width, height (height includes dropdown)
        hWnd,
        NULL,
        (HINSTANCE)GetWindowLongPtr(hWnd, GWLP_HINSTANCE),
        NULL
    );

    // Edit box (below ComboBox)
    hEdit = CreateWindowEx(
        WS_EX_CLIENTEDGE,
        TEXT("EDIT"),
        TEXT("If other..."),
        WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
        padding + inputWidth + 10, padding + labelHeight + controlHeight + 12,     // x, y
        comboWidth, controlHeight,     // width, height
        hWnd,
        (HMENU)TYPE_INPUT,
        GetModuleHandle(NULL),
        NULL
    );

    // CONVERT button (centered, below all controls)
    int convertBtnX = (winWidth - buttonWidth) / 2;
    int convertBtnY = 150; //winHeight - buttonHeight - padding;

    HWND convertButton = CreateWindowEx(
        0,
        L"BUTTON",
        L"Convert",
        WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
        convertBtnX, convertBtnY, // x, y (centered horizontally)
        buttonWidth, buttonHeight,
        hWnd,
        (HMENU)CONVERT_BUTTON,
        (HINSTANCE)GetWindowLongPtr(hWnd, GWLP_HINSTANCE),
        NULL
    );

    // Fill ComboBox
    for (int i = 0; i < ARRAYSIZE(LfileExtensions); i++) {
        SendMessage(typeSelector, CB_ADDSTRING, (WPARAM)0, (LPARAM) LfileExtensions[i]);
    }
    SendMessage(typeSelector, CB_SETCURSEL, (WPARAM)2, (LPARAM)0);

    // Set font and modern visuals for all controls and main window
    SetModernFont(hWnd);

    SendMessage(fileLocation, WM_SETFONT, (WPARAM)GetStockObject(DEFAULT_GUI_FONT), TRUE);
    SendMessage(typeLabel, WM_SETFONT, (WPARAM)GetStockObject(DEFAULT_GUI_FONT), TRUE);
    SendMessage(typeSelector, WM_SETFONT, (WPARAM)GetStockObject(DEFAULT_GUI_FONT), TRUE);
    SendMessage(hEdit, WM_SETFONT, (WPARAM)GetStockObject(DEFAULT_GUI_FONT), TRUE);

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
