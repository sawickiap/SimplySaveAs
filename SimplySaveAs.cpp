#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>
#include <shobjidl.h>
#include <atlbase.h>

#include <string>
#include <vector>
#include <stdexcept>

using std::wstring;
using std::vector;

#include <cstdint>
#include <cassert>
#include <cstdlib>
#include <cctype>
#include <cstdarg>

static const wchar_t* const APP_TITLE = L"SimplySaveAs";
static const wchar_t* const MAIN_MESSAGE =
    L"SimplySaveAs 1.0\n"
    L"Copyright (C) Adam Sawicki. License: MIT.\n"
    L"https://github.com/sawickiap/SimplySaveAs\n"
    L"\n"
    L"To use this program you must run it with path to source file specified as command line parameter.";

enum class Result : int32_t
{
    Exception = -5,
    CopyError = -4,
    DstInvalid = -3,
    SrcInvalid = -2,
    CmdLineInvalid = -1,
    Success = 0,
    Cancelled = 1,
};

#define CHECK_HR(expr) if(HRESULT hr_ = (expr); FAILED(hr_)) { assert(0 && #expr); }

class CommandLineParser
{
public:
    CommandLineParser(const wchar_t* cmdLine) :
        m_CmdLine{cmdLine},
        m_CmdLineLen{wcslen(cmdLine)},
        m_Index{0}
    {
    }

    bool GetNextParameter(wstring& out);

private:
    const wchar_t* const m_CmdLine;
    const size_t m_CmdLineLen;
    size_t m_Index;
};

bool CommandLineParser::GetNextParameter(wstring& out)
{
    out.clear();
    while(m_Index < m_CmdLineLen && isspace(m_CmdLine[m_Index]))
        ++m_Index;
    if(m_Index == m_CmdLineLen)
        return false;
    if(m_CmdLine[m_Index] == L'"')
    {
        ++m_Index;
        out += m_CmdLine[m_Index++];
        while(m_Index < m_CmdLineLen && m_CmdLine[m_Index] != L'"')
            out += m_CmdLine[m_Index++];
        if(m_Index < m_CmdLineLen && m_CmdLine[m_Index] == L'"')
            ++m_Index;
    }
    else
    {
        out += m_CmdLine[m_Index++];
        while(m_Index < m_CmdLineLen && !isspace(m_CmdLine[m_Index]))
            out += m_CmdLine[m_Index++];
    }
    return true;
}

static wstring ExtractFileName(const wstring& path)
{
    const size_t lastSlashPos = path.find_last_of(L"/\\");
    if(lastSlashPos == wstring::npos)
        return path;
    else
        return path.substr(lastSlashPos + 1);
}

static void ShowMessage(bool error, const wchar_t* format, ...)
{
    va_list argList;
    va_start(argList, format);
    const size_t dstLen = (size_t)::_vscwprintf(format, argList);
    vector<wchar_t> buf(dstLen + 1);
    vswprintf_s(buf.data(), buf.size(), format, argList);

    const DWORD iconType = error ? MB_ICONERROR : MB_ICONINFORMATION;
    MessageBox(NULL, buf.data(), APP_TITLE, MB_OK | iconType);

    va_end(argList);
}

static bool Dialog(wstring& outPath, const wchar_t* fileName)
{
    CComPtr<IFileSaveDialog> dialog;
    CHECK_HR(CoCreateInstance(CLSID_FileSaveDialog, NULL, CLSCTX_ALL, IID_IFileSaveDialog, (void**)&dialog));

    FILEOPENDIALOGOPTIONS options = 0;
    CHECK_HR(dialog->GetOptions(&options));
    options |= FOS_FORCEFILESYSTEM;
    CHECK_HR(dialog->SetOptions(options));
    CHECK_HR(dialog->SetFileName(fileName));

    HRESULT hr = dialog->Show(NULL);
    if (SUCCEEDED(hr))
    {
        CComPtr<IShellItem> pItem;
        hr = dialog->GetResult(&pItem);
        if(SUCCEEDED(hr))
        {
            PWSTR fileSysPath = NULL;
            hr = pItem->GetDisplayName(SIGDN_FILESYSPATH, &fileSysPath);
            if(SUCCEEDED(hr))
            {
                outPath = fileSysPath;
                CoTaskMemFree(fileSysPath);
            }
            return true;
        }

    }
    return false;
}

static Result Copy(const wstring& dstPath, const wstring& srcPath)
{
    if(_wcsicmp(dstPath.c_str(), srcPath.c_str()) == 0)
    {
        ShowMessage(true, L"Cannot copy to the same location as source path \"%s\".", srcPath.c_str());
        return Result::DstInvalid;
    }

    BOOL cancel = FALSE;
    const BOOL ok = CopyFileEx(
        srcPath.c_str(), // lpExistingFileName
        dstPath.c_str(), // lpNewFileName
        NULL, // lpProgressRoutine
        NULL, // lpData
        &cancel, // pbCancel
        0); // dwCopyFlags
    if(ok)
    {
        return Result::Success;
    }
    else
    {
        ShowMessage(true, L"Couldn't copy \"%s\" to \"%s\".", srcPath.c_str(), dstPath.c_str());
        return Result::CopyError;
    }
}

static bool ParseCommandLine(wstring& outSrcPath, const wchar_t* cmdLine)
{
    CommandLineParser parser{cmdLine};
    bool success = false;
    if(parser.GetNextParameter(outSrcPath))
    {
        wstring dummy;
        success = !parser.GetNextParameter(dummy);
    }

    if(!success)
    {
        ShowMessage(false, MAIN_MESSAGE);
    }
    return success;
}

static int Main2(const wchar_t* cmdLine)
{
    CHECK_HR(CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE));
    Result result = Result::Success;

    wstring srcPath;
    if(ParseCommandLine(srcPath, cmdLine))
    {
        const DWORD srcAttr = GetFileAttributes(srcPath.c_str());
        if(srcAttr == INVALID_FILE_ATTRIBUTES)
        {
            ShowMessage(true, L"Source file \"%s\" doesn't exist.", srcPath.c_str());
            result = Result::SrcInvalid;
        }
        else if(srcAttr & FILE_ATTRIBUTE_DIRECTORY)
        {
            ShowMessage(true, L"Source \"%s\" is a directory.", srcPath.c_str());
            result = Result::SrcInvalid;
        }
        else
        {
            wstring srcFileName = ExtractFileName(srcPath);
            wstring dstPath;
            if(Dialog(dstPath, srcFileName.c_str()))
            {
                result = Copy(dstPath, srcPath);
            }
            else
            {
                result = Result::Cancelled;
            }
        }
    }
    else
    {
        result = Result::CmdLineInvalid;
    }
    CoUninitialize();
    return (int)result;
}

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR cmdLine, int cmdShow)
{
    try
    {
        return Main2(cmdLine);
    }
    catch(const std::exception& ex)
    {
        ShowMessage(true, L"%hs", ex.what());
        return (int)Result::Exception;
    }
    catch(...)
    {
        ShowMessage(true, L"Unknown error.");
        return (int)Result::Exception;
    }
}
