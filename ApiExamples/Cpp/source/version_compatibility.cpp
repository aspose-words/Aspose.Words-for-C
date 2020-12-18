#include "version_info/version_info.h"
#include "version_info/version_defines.h"
#include <cstdlib>
#include <cstring>
#include <iostream>
#include <sstream>

#ifdef __linux__
#include <codecvt>
#include <dlfcn.h>
#include <locale>
#include <string>

#define USED_ATTRIBUTE __attribute__((used))
#else
#include "windows.h"
#define USED_ATTRIBUTE
#endif

#define VERSION_MISMATCH_ACTION_STDERR 1
#define VERSION_MISMATCH_ACTION_STDOUT 2
#define VERSION_MISMATCH_ACTION_EXIT 3
#define VERSION_MISMATCH_ACTION_SILENT 4

#ifdef VERSION_MISMATCH_ACTION
#undef VERSION_MISMATCH_ACTION
#endif
#define VERSION_MISMATCH_ACTION VERSION_MISMATCH_ACTION_STDERR

namespace {

class ModulesCompatibilityChecker
{
public:
    ModulesCompatibilityChecker()
    {
#ifdef VERSION_COMPATIBILITY_TEST_MODE
        return;
#else
        CheckVersions();
#endif
    }

    static void CheckVersions()
    {
        CodePorting::Native::Cs2Cpp::VersionInfo asposecpplib_version_info;
        CodePorting::Native::Cs2Cpp::_getVersionInfo(asposecpplib_version_info);

        CodePorting::Native::Cs2Cpp::VersionInfo product_version_info = {0};
        FillExpectedVersionInformation(product_version_info);

        if (asposecpplib_version_info != product_version_info)
        {
#if VERSION_MISMATCH_ACTION != VERSION_MISMATCH_ACTION_EXIT
            CodePorting::Native::Cs2Cpp::MismatchedVersionInfo mvi;
            GetModuleName(mvi.module_name);
            mvi.product_version_info = product_version_info;

            CodePorting::Native::Cs2Cpp::_registerVersionMismatch(mvi);
#endif

#if VERSION_MISMATCH_ACTION == VERSION_MISMATCH_ACTION_STDOUT
            PrintErrorMessage(std::cout, product_version_info, asposecpplib_version_info);
#elif VERSION_MISMATCH_ACTION == VERSION_MISMATCH_ACTION_EXIT
            PrintErrorMessage(std::cerr, product_version_info, asposecpplib_version_info);
#ifdef VERSION_COMPATIBILITY_TEST_MODE
            return; // Don't exit if we are in test mode.
#else
            exit(EXIT_FAILURE);
#endif
#elif VERSION_MISMATCH_ACTION == VERSION_MISMATCH_ACTION_SILENT
            // NOOP
#else
            // Default is VERSION_MISMATCH_ACTION_STDERR
            PrintErrorMessage(std::cerr, product_version_info, asposecpplib_version_info);
#endif
        }
    }

    static void GetModuleName(wchar_t (&module_file_name)[MODULE_NAME_BUFFER_SIZE])
    {
        module_file_name[0] = L'\0'; // Assume failure.
#ifdef _WIN32
        HMODULE h_module = 0;
        BOOL success = GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                                          reinterpret_cast<LPCWSTR>(&GetModuleName), &h_module);

        if (success)
            GetModuleFileNameW(h_module, module_file_name, MODULE_NAME_BUFFER_SIZE);

#elif defined __linux__
        Dl_info dl_info;
        int res = dladdr(reinterpret_cast<void*>(&GetModuleName), &dl_info);

        if (res > 0)
        {
            std::string utf8_module_name = dl_info.dli_fname;
            std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>> converter;
            std::wstring utf16_module_name = converter.from_bytes(utf8_module_name.c_str());

            wcsncpy(module_file_name, utf16_module_name.c_str(), MODULE_NAME_BUFFER_SIZE);
            module_file_name[MODULE_NAME_BUFFER_SIZE - 1] = L'\0';
        }
#else
        // TODO: Generate compilation error.
#endif
    }

    static void PrintErrorMessage(std::ostream& destination, const CodePorting::Native::Cs2Cpp::VersionInfo& product_version_info,
                                  const CodePorting::Native::Cs2Cpp::VersionInfo& asposecpplib_version_info)
    {
        char encoded_product_version_info[CodePorting::Native::Cs2Cpp::VersionInfo::ENCODED_VERSION_INFO_SIZE];
        product_version_info.encode(encoded_product_version_info);

        char encoded_asposecpplib_version_info[CodePorting::Native::Cs2Cpp::VersionInfo::ENCODED_VERSION_INFO_SIZE];
        asposecpplib_version_info.encode(encoded_asposecpplib_version_info);

        destination << "Modules versions mismatch!" << std::endl
                    << "ProductVersionInfo:      " << static_cast<const char*>(encoded_product_version_info) << std::endl
                    << "CodePortingNativeCs2CppVersionInfo: " << static_cast<const char*>(encoded_asposecpplib_version_info) << std::endl;
    }

    static void FillExpectedVersionInformation(CodePorting::Native::Cs2Cpp::VersionInfo& product_version_info)
    {
        product_version_info.asposecpplib_version_major = ASPOSECPPLIB_VERSION_MAJOR;
        product_version_info.asposecpplib_version_minor = ASPOSECPPLIB_VERSION_MINOR;
        product_version_info.asposecpplib_version_revision = ASPOSECPPLIB_VERSION_REVISION;
        product_version_info.asposecpplib_version_build = ASPOSECPPLIB_VERSION_BUILD;

        product_version_info.product_version_major = PRODUCT_VERSION_MAJOR;
        product_version_info.product_version_minor = PRODUCT_VERSION_MINOR;
        product_version_info.product_version_revision = PRODUCT_VERSION_REVISION;
        product_version_info.product_version_build = PRODUCT_VERSION_BUILD;

        strncpy_s(product_version_info.asposecpplib_revision, ASPOSECPPLIB_REVISION, GIT_REVISION_BUFFER_SIZE - 1);

#ifdef __DBG_FOR_EACH_MEMBER
        product_version_info.defined___DBG_FOR_EACH_MEMEBR = true;
#else
        product_version_info.defined___DBG_FOR_EACH_MEMEBR = false;
#endif

#ifdef __DBG_GARBAGE_COLLECTION
        product_version_info.defined___DBG_GARBAGE_COLLECTION = true;
#else
        product_version_info.defined___DBG_GARBAGE_COLLECTION = false;
#endif

#ifdef __DBG_TOOLS
        product_version_info.defined___DBG_TOOLS = true;
#else
        product_version_info.defined___DBG_TOOLS = false;
#endif

#ifdef ASPOSE_COMPARE_BY_REF
        product_version_info.defined_ASPOSE_COMPARE_BY_REF = true;
#else
        product_version_info.defined_ASPOSE_COMPARE_BY_REF = false;
#endif

#ifdef ASPOSE_THREADSAFE_DELEGATES
        product_version_info.defined_ASPOSE_THREADSAFE_DELEGATES = true;
#else
        product_version_info.defined_ASPOSE_THREADSAFE_DELEGATES = false;
#endif

#ifdef ASPOSE_NO_ATOMIC_REFCOUNT
        product_version_info.defined_ASPOSE_NO_ATOMIC_REFCOUNT = true;
#else
        product_version_info.defined_ASPOSE_NO_ATOMIC_REFCOUNT = false;
#endif

#ifdef ENABLE_EXTERNAL_REFCOUNT
        product_version_info.defined_ENABLE_EXTERNAL_REFCOUNT = true;
#else
        product_version_info.defined_ENABLE_EXTERNAL_REFCOUNT = false;
#endif

#ifdef DISABLE_ASTRAL
        product_version_info.defined_DISABLE_ASTRAL = true;
#else
        product_version_info.defined_DISABLE_ASTRAL = false;
#endif

#ifdef CALL_DISPOSE
        product_version_info.defined_CALL_DISPOSE = true;
#else
        product_version_info.defined_CALL_DISPOSE = false;
#endif

#ifdef ENABLE_CYCLES_DETECTION_EXT
        product_version_info.defined_ENABLE_CYCLES_DETECTION_EXT = true;
#else
        product_version_info.defined_ENABLE_CYCLES_DETECTION_EXT = false;
#endif

#ifdef ENABLE_MAKE_OBJECT_LEAKAGE_DETECTION
        product_version_info.defined_ENABLE_MAKE_OBJECT_LEAKAGE_DETECTION = true;
#else
        product_version_info.defined_ENABLE_MAKE_OBJECT_LEAKAGE_DETECTION = false;
#endif

#ifdef ENABLE_CYCLES_DETECTION
        product_version_info.defined_ENABLE_CYCLES_DETECTION = true;
#else
        product_version_info.defined_ENABLE_CYCLES_DETECTION = false;
#endif

#ifdef SHOW_DISPOSE_GUARD_MESSAGE
        product_version_info.defined_SHOW_DISPOSE_GUARD_MESSAGE = true;
#else
        product_version_info.defined_SHOW_DISPOSE_GUARD_MESSAGE = false;
#endif
    }
};

ModulesCompatibilityChecker USED_ATTRIBUTE a;

} // namespace
