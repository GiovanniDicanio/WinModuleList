//////////////////////////////////////////////////////////////////////////
//
// Module Enumerator for a Windows Process (console application)
// -------------------------------------------------------------
// List all the modules in a Windows process, given the process ID.
//
// Copyright (c) 2024 by Giovanni Dicanio
//
//////////////////////////////////////////////////////////////////////////


//------------------------------------------------------------------------
//                              Includes
//------------------------------------------------------------------------

#include <Windows.h>    // Main Windows Platform SDK Header
#include <Shlwapi.h>    // StrToInt64ExW
#include <TlHelp32.h>   // Tool Help Library

#include <exception>    // std::exception
#include <iostream>     // Console output
#include <stdexcept>    // std::runtime_error
#include <string>       // std::wstring
#include <vector>       // std::vector

#pragma comment(lib, "shlwapi.lib")


//------------------------------------------------------------------------
// Helper class: Simplify our life a bit with a simple scoped-wrapper
// around raw Win32 HANDLE.
// Will *automatically* invoke ::CloseHandle when it goes out of scope.
//------------------------------------------------------------------------
class ScopedHandle
{
public:

    // Get ownership of the input handle
    explicit ScopedHandle(HANDLE h) noexcept
        : m_handle{ h }
    {}

    // Automatically close the handle
    ~ScopedHandle() noexcept
    {
        ::CloseHandle(m_handle);
        m_handle = INVALID_HANDLE_VALUE;
    }

    // Access the wrapped raw handle
    HANDLE Get() const noexcept
    {
        return m_handle;
    }


    //
    // Ban copy and move
    //
private:
    ScopedHandle(ScopedHandle const&) = delete;
    ScopedHandle& operator=(ScopedHandle const&) = delete;

    ScopedHandle(ScopedHandle&&) = delete;
    ScopedHandle& operator=(ScopedHandle&&) = delete;


    //
    // Data Members
    //
private:
    HANDLE m_handle;
};


//------------------------------------------------------------------------
// Exception class representing a Windows API error
// (e.g. error code as returned by GetLastError)
//------------------------------------------------------------------------
class Win32Error : public std::runtime_error
{
public:
    Win32Error(const char* errorMessage, DWORD errorCode)
        : std::runtime_error{ errorMessage }
        , m_errorCode{ errorCode }
    {}

    DWORD GetErrorCode() const noexcept
    {
        return m_errorCode;
    }

private:
    DWORD m_errorCode;
};

void ThrowLastWin32Error(const char* message)
{
    const DWORD errorCode = ::GetLastError();
    throw Win32Error(message, errorCode);
}


//------------------------------------------------------------------------
// Info for a module in the module list.
// Store module name and size; but could add also other fields
// from the MODULEENTRY32 Win32 structure.
//------------------------------------------------------------------------
struct ModuleInfo
{
    std::wstring Name;
    DWORD Size;

    ModuleInfo(std::wstring const& name, DWORD size)
        : Name{ name }, Size{ size }
    {}
};

//------------------------------------------------------------------------
// Enumerate the modules in a given process
//------------------------------------------------------------------------
std::vector<ModuleInfo> GetModuleListInProcess(DWORD processID)
{
    std::vector<ModuleInfo> moduleList;

    // Create module snapshot for the enumeration process
    HANDLE hEnum = ::CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, processID);
    if (hEnum == INVALID_HANDLE_VALUE)
    {
        ThrowLastWin32Error("Cannot init the module enumeration process (CreateToolhelp32Snapshot failed).");
    }

    // Safely wrap the raw HANDLE in a "RAII" wrapper, such that it will be *automatically* released
    // via a call to CloseHandle also in case of exceptions thrown
    ScopedHandle enumHandle{ hEnum };

    MODULEENTRY32 moduleEntry = { sizeof(moduleEntry) };

    // Start the enumeration process
    BOOL continueEnumeration = ::Module32First(enumHandle.Get(), &moduleEntry);
    while (continueEnumeration)
    {
        // Extract the pieces of information we need from the MODULEENTRY32 structure,
        // pack them in a proper C++ structure, and add that to the moduleList vector
        moduleList.push_back(ModuleInfo{ moduleEntry.szModule, moduleEntry.modBaseSize });

        // Move to the next module (if any)
        continueEnumeration = ::Module32Next(enumHandle.Get(), &moduleEntry);
    }
    // Make sure that the enumeration has been terminated because there are no more modules
    // in the list (in this case, GetLastError will return ERROR_NO_MORE_FILES as per MSDN doc).
    const DWORD lastError = ::GetLastError();
    if (lastError != ERROR_NO_MORE_FILES)
    {
        ThrowLastWin32Error("Unexpected termination of the module enumeration process.");
    }

    return moduleList;
}


//------------------------------------------------------------------------
// Console Application Entry-Point
//------------------------------------------------------------------------
int wmain(int argc, wchar_t* argv[])
{
    constexpr int kExitOk = 0;
    constexpr int kExitError = 1;

    try
    {
        std::wcout << L"\n *** Enumerate Modules in a Process *** \n";
        std::wcout << L"          by Giovanni Dicanio \n\n";

        //
        // Parse the command line parameter(s):
        // Here we accept the process ID for the module enumeration as the only parameter
        //

        if (argc != 2)
        {
            std::wcout << L"Please pass the ID of the process to enumerate as the only parameter.\n";
            return kExitError;
        }

        LONGLONG llResult = 0;
        if (!::StrToInt64ExW(argv[1], STIF_DEFAULT, &llResult))
        {
            std::wcout << L"Please pass the ID of the process to enumerate as an integer in base 10.\n";
            return kExitError;
        }
        const DWORD processID = static_cast<DWORD>(llResult);


        //
        // Enumerate the modules in the given process
        //

        std::vector<ModuleInfo> moduleList = GetModuleListInProcess(processID);


        //
        // Print out the module list
        //

        std::wcout << L" Module List for Process ID = " << processID << L"\n";
        std::wcout << L" ========================================\n\n";
        for (const auto& moduleInfo : moduleList)
        {
            std::wcout << L" - " << moduleInfo.Name << L"  (" << moduleInfo.Size << L" bytes)\n";
        }

        return kExitOk;
    }
    //
    // Handle errors
    //
    catch (Win32Error const& e)
    {
        std::wcout << L"*** ERROR: " << e.what() << L'\n';
        std::wcout << L"    (Error code: " << e.GetErrorCode() << L")\n";
        return kExitError;
    }
    catch (std::exception const& e)
    {
        std::wcout << L"*** ERROR: " << e.what() << L'\n';
        return kExitError;
    }
}

//////////////////////////////////////////////////////////////////////////
