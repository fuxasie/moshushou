#include <windows.h>
#include <setupapi.h>
#include <newdev.h>
#include <cfgmgr32.h>
#include <wincrypt.h>
#include <winternl.h>
#include <io.h>
#include <fcntl.h>

#include <filesystem>
#include <fstream>
#include <iostream>
#include <string>
#include <vector>

namespace
{
    constexpr wchar_t HardwareId[] = L"Root\\MoshushouVirtualHid";
    constexpr DWORD MinimumBuild = 18362;

    struct DeviceState
    {
        bool Found = false;
        bool Started = false;
        ULONG Problem = 0;
        std::wstring InstanceId;
    };

    class DeviceInfoSet
    {
    public:
        explicit DeviceInfoSet(HDEVINFO value = INVALID_HANDLE_VALUE) : value_(value) {}
        ~DeviceInfoSet()
        {
            if (value_ != INVALID_HANDLE_VALUE)
            {
                SetupDiDestroyDeviceInfoList(value_);
            }
        }

        DeviceInfoSet(const DeviceInfoSet&) = delete;
        DeviceInfoSet& operator=(const DeviceInfoSet&) = delete;
        HDEVINFO Get() const { return value_; }
        bool IsValid() const { return value_ != INVALID_HANDLE_VALUE; }

    private:
        HDEVINFO value_;
    };

    std::wstring ErrorMessage(DWORD error)
    {
        wchar_t* buffer = nullptr;
        DWORD length = FormatMessageW(
            FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
            nullptr,
            error,
            0,
            reinterpret_cast<wchar_t*>(&buffer),
            0,
            nullptr);

        std::wstring message = length && buffer ? std::wstring(buffer, length) : L"Unknown error";
        if (buffer)
        {
            LocalFree(buffer);
        }
        while (!message.empty() && (message.back() == L'\r' || message.back() == L'\n'))
        {
            message.pop_back();
        }
        return message;
    }

    bool EqualsInsensitive(const wchar_t* left, const wchar_t* right)
    {
        return left && right && _wcsicmp(left, right) == 0;
    }

    bool IsAdministrator()
    {
        HANDLE token = nullptr;
        if (!OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &token))
        {
            return false;
        }

        TOKEN_ELEVATION elevation{};
        DWORD returned = 0;
        bool elevated = GetTokenInformation(
            token,
            TokenElevation,
            &elevation,
            sizeof(elevation),
            &returned) && elevation.TokenIsElevated != 0;
        CloseHandle(token);
        return elevated;
    }

    bool IsSupportedSystem(DWORD& build)
    {
        SYSTEM_INFO systemInfo{};
        GetNativeSystemInfo(&systemInfo);
        if (systemInfo.wProcessorArchitecture != PROCESSOR_ARCHITECTURE_AMD64)
        {
            std::wcerr << L"Only native x64 Windows is supported.\n";
            return false;
        }

        using RtlGetVersionFunction = LONG(WINAPI*)(PRTL_OSVERSIONINFOW);
        auto rtlGetVersion = reinterpret_cast<RtlGetVersionFunction>(
            GetProcAddress(GetModuleHandleW(L"ntdll.dll"), "RtlGetVersion"));
        if (!rtlGetVersion)
        {
            std::wcerr << L"RtlGetVersion is unavailable.\n";
            return false;
        }

        RTL_OSVERSIONINFOW version{};
        version.dwOSVersionInfoSize = sizeof(version);
        if (rtlGetVersion(&version) != 0)
        {
            std::wcerr << L"Unable to read the Windows version.\n";
            return false;
        }

        build = version.dwBuildNumber;
        if (version.dwMajorVersion < 10 || build < MinimumBuild)
        {
            std::wcerr << L"Windows build " << build
                       << L" is unsupported; build 18362 or newer is required.\n";
            return false;
        }
        return true;
    }

    bool HasMatchingHardwareId(HDEVINFO set, SP_DEVINFO_DATA& device)
    {
        DWORD type = 0;
        DWORD required = 0;
        SetupDiGetDeviceRegistryPropertyW(
            set,
            &device,
            SPDRP_HARDWAREID,
            &type,
            nullptr,
            0,
            &required);
        if (GetLastError() != ERROR_INSUFFICIENT_BUFFER || required < sizeof(wchar_t))
        {
            return false;
        }

        std::vector<BYTE> buffer(required + sizeof(wchar_t) * 2, 0);
        if (!SetupDiGetDeviceRegistryPropertyW(
                set,
                &device,
                SPDRP_HARDWAREID,
                &type,
                buffer.data(),
                static_cast<DWORD>(buffer.size()),
                nullptr))
        {
            return false;
        }

        const wchar_t* current = reinterpret_cast<const wchar_t*>(buffer.data());
        while (*current)
        {
            if (EqualsInsensitive(current, HardwareId))
            {
                return true;
            }
            current += wcslen(current) + 1;
        }
        return false;
    }

    DeviceState QueryDeviceState()
    {
        DeviceState result;
        DeviceInfoSet devices(SetupDiGetClassDevsW(
            nullptr,
            nullptr,
            nullptr,
            DIGCF_ALLCLASSES | DIGCF_PRESENT));
        if (!devices.IsValid())
        {
            return result;
        }

        for (DWORD index = 0;; ++index)
        {
            SP_DEVINFO_DATA device{};
            device.cbSize = sizeof(device);
            if (!SetupDiEnumDeviceInfo(devices.Get(), index, &device))
            {
                if (GetLastError() == ERROR_NO_MORE_ITEMS)
                {
                    break;
                }
                continue;
            }

            if (!HasMatchingHardwareId(devices.Get(), device))
            {
                continue;
            }

            result.Found = true;
            wchar_t instanceId[MAX_DEVICE_ID_LEN]{};
            if (SetupDiGetDeviceInstanceIdW(
                    devices.Get(),
                    &device,
                    instanceId,
                    ARRAYSIZE(instanceId),
                    nullptr))
            {
                result.InstanceId = instanceId;
            }

            ULONG status = 0;
            ULONG problem = 0;
            if (CM_Get_DevNode_Status(&status, &problem, device.DevInst, 0) == CR_SUCCESS)
            {
                result.Problem = problem;
                result.Started = (status & DN_STARTED) != 0 && problem == 0;
            }
            return result;
        }
        return result;
    }

    bool AddCertificateToStore(PCCERT_CONTEXT certificate, const wchar_t* storeName)
    {
        HCERTSTORE store = CertOpenStore(
            CERT_STORE_PROV_SYSTEM_W,
            0,
            0,
            CERT_SYSTEM_STORE_LOCAL_MACHINE | CERT_STORE_OPEN_EXISTING_FLAG,
            storeName);
        if (!store)
        {
            std::wcerr << L"CertOpenStore(" << storeName << L") failed: "
                       << ErrorMessage(GetLastError()) << L"\n";
            return false;
        }

        BOOL added = CertAddCertificateContextToStore(
            store,
            certificate,
            CERT_STORE_ADD_REPLACE_EXISTING,
            nullptr);
        DWORD error = added ? ERROR_SUCCESS : GetLastError();
        CertCloseStore(store, 0);
        if (!added)
        {
            std::wcerr << L"Adding the certificate to " << storeName << L" failed: "
                       << ErrorMessage(error) << L"\n";
            return false;
        }
        return true;
    }

    bool InstallCertificate(const std::filesystem::path& certificatePath)
    {
        std::ifstream stream(certificatePath, std::ios::binary | std::ios::ate);
        if (!stream)
        {
            std::wcerr << L"Certificate file is unavailable: " << certificatePath.c_str() << L"\n";
            return false;
        }

        std::streamsize size = stream.tellg();
        if (size <= 0 || size > 1024 * 1024)
        {
            std::wcerr << L"Certificate file size is invalid.\n";
            return false;
        }
        stream.seekg(0, std::ios::beg);
        std::vector<BYTE> bytes(static_cast<size_t>(size));
        if (!stream.read(reinterpret_cast<char*>(bytes.data()), size))
        {
            std::wcerr << L"Reading the certificate failed.\n";
            return false;
        }

        PCCERT_CONTEXT certificate = CertCreateCertificateContext(
            X509_ASN_ENCODING | PKCS_7_ASN_ENCODING,
            bytes.data(),
            static_cast<DWORD>(bytes.size()));
        if (!certificate)
        {
            std::wcerr << L"Invalid certificate file: " << ErrorMessage(GetLastError()) << L"\n";
            return false;
        }

        bool success = AddCertificateToStore(certificate, L"ROOT") &&
                       AddCertificateToStore(certificate, L"TrustedPublisher");
        CertFreeCertificateContext(certificate);
        return success;
    }

    bool CreateRootDevice(const std::filesystem::path& infPath)
    {
        GUID classGuid{};
        wchar_t className[MAX_CLASS_NAME_LEN]{};
        if (!SetupDiGetINFClassW(
                infPath.c_str(),
                &classGuid,
                className,
                ARRAYSIZE(className),
                nullptr))
        {
            std::wcerr << L"SetupDiGetINFClass failed: " << ErrorMessage(GetLastError()) << L"\n";
            return false;
        }

        DeviceInfoSet devices(SetupDiCreateDeviceInfoList(&classGuid, nullptr));
        if (!devices.IsValid())
        {
            std::wcerr << L"SetupDiCreateDeviceInfoList failed: "
                       << ErrorMessage(GetLastError()) << L"\n";
            return false;
        }

        SP_DEVINFO_DATA device{};
        device.cbSize = sizeof(device);
        if (!SetupDiCreateDeviceInfoW(
                devices.Get(),
                className,
                &classGuid,
                L"Moshushou Virtual Keyboard and Mouse",
                nullptr,
                DICD_GENERATE_ID,
                &device))
        {
            std::wcerr << L"SetupDiCreateDeviceInfo failed: " << ErrorMessage(GetLastError()) << L"\n";
            return false;
        }

        std::vector<wchar_t> hardwareIds(wcslen(HardwareId) + 2, L'\0');
        wcscpy_s(hardwareIds.data(), hardwareIds.size(), HardwareId);
        if (!SetupDiSetDeviceRegistryPropertyW(
                devices.Get(),
                &device,
                SPDRP_HARDWAREID,
                reinterpret_cast<const BYTE*>(hardwareIds.data()),
                static_cast<DWORD>(hardwareIds.size() * sizeof(wchar_t))))
        {
            std::wcerr << L"Setting the hardware ID failed: " << ErrorMessage(GetLastError()) << L"\n";
            return false;
        }

        if (!SetupDiCallClassInstaller(DIF_REGISTERDEVICE, devices.Get(), &device))
        {
            std::wcerr << L"Registering the root device failed: "
                       << ErrorMessage(GetLastError()) << L"\n";
            return false;
        }
        return true;
    }

    bool RemoveMatchingDevices(BOOL* rebootRequired = nullptr, size_t* removedCount = nullptr)
    {
        bool success = true;
        BOOL anyRebootRequired = FALSE;
        size_t removed = 0;
        DeviceInfoSet devices(SetupDiGetClassDevsW(
            nullptr,
            nullptr,
            nullptr,
            DIGCF_ALLCLASSES));
        if (!devices.IsValid())
        {
            return false;
        }

        for (DWORD index = 0;; ++index)
        {
            SP_DEVINFO_DATA device{};
            device.cbSize = sizeof(device);
            if (!SetupDiEnumDeviceInfo(devices.Get(), index, &device))
            {
                if (GetLastError() == ERROR_NO_MORE_ITEMS)
                {
                    break;
                }
                continue;
            }
            if (!HasMatchingHardwareId(devices.Get(), device))
            {
                continue;
            }

            wchar_t instanceId[MAX_DEVICE_ID_LEN]{};
            SetupDiGetDeviceInstanceIdW(
                devices.Get(),
                &device,
                instanceId,
                ARRAYSIZE(instanceId),
                nullptr);

            BOOL deviceRebootRequired = FALSE;
            if (!DiUninstallDevice(
                    nullptr,
                    devices.Get(),
                    &device,
                    0,
                    &deviceRebootRequired))
            {
                DWORD error = GetLastError();
                std::wcerr << L"DiUninstallDevice failed for " << instanceId << L": "
                           << ErrorMessage(error) << L" (" << error << L")\n";
                success = false;
            }
            else
            {
                ++removed;
                anyRebootRequired = anyRebootRequired || deviceRebootRequired;
            }
        }

        if (rebootRequired)
        {
            *rebootRequired = anyRebootRequired;
        }
        if (removedCount)
        {
            *removedCount = removed;
        }
        return success;
    }

    int PrintStatus()
    {
        DeviceState state = QueryDeviceState();
        if (!state.Found)
        {
            std::wcout << L"NOT_INSTALLED\n";
            return 2;
        }
        if (!state.Started)
        {
            std::wcout << L"INSTALLED_NOT_STARTED problem=" << state.Problem
                       << L" instance=" << state.InstanceId << L"\n";
            return 3;
        }

        std::wcout << L"STARTED instance=" << state.InstanceId << L"\n";
        return 0;
    }

    int InstallDriver(
        const std::filesystem::path& infPath,
        const std::filesystem::path& certificatePath)
    {
        if (!IsAdministrator())
        {
            std::wcerr << L"Administrator privileges are required.\n";
            return 5;
        }

        DWORD build = 0;
        if (!IsSupportedSystem(build))
        {
            return 4;
        }
        if (!std::filesystem::is_regular_file(infPath))
        {
            std::wcerr << L"INF file is unavailable: " << infPath.c_str() << L"\n";
            return 5;
        }
        if (!certificatePath.empty() && !InstallCertificate(certificatePath))
        {
            return 1;
        }

        DeviceState before = QueryDeviceState();
        bool created = false;
        if (!before.Found)
        {
            if (!CreateRootDevice(infPath))
            {
                return 1;
            }
            created = true;
        }

        BOOL rebootRequired = FALSE;
        if (!UpdateDriverForPlugAndPlayDevicesW(
                nullptr,
                HardwareId,
                infPath.c_str(),
                INSTALLFLAG_FORCE,
                &rebootRequired))
        {
            DWORD error = GetLastError();
            std::wcerr << L"UpdateDriverForPlugAndPlayDevices failed: "
                       << ErrorMessage(error) << L" (" << error << L")\n";
            if (created)
            {
                RemoveMatchingDevices();
            }
            return 1;
        }

        for (int attempt = 0; attempt < 50; ++attempt)
        {
            DeviceState state = QueryDeviceState();
            if (state.Found && state.Started)
            {
                std::wcout << L"INSTALL_SUCCESS build=" << build
                           << L" instance=" << state.InstanceId
                           << L" reboot=" << (rebootRequired ? L"true" : L"false") << L"\n";
                return rebootRequired ? 3010 : 0;
            }
            Sleep(200);
        }

        DeviceState state = QueryDeviceState();
        std::wcerr << L"The device was installed but did not start. problem="
                   << state.Problem << L"\n";
        return 3;
    }

    int UninstallDriver(const std::filesystem::path& infPath)
    {
        if (!IsAdministrator())
        {
            std::wcerr << L"Administrator privileges are required.\n";
            return 5;
        }

        DWORD build = 0;
        if (!IsSupportedSystem(build))
        {
            return 4;
        }
        if (!std::filesystem::is_regular_file(infPath))
        {
            std::wcerr << L"INF file is unavailable: " << infPath.c_str() << L"\n";
            return 5;
        }

        BOOL driverRebootRequired = FALSE;
        bool driverPackageRemoved = DiUninstallDriverW(
            nullptr,
            infPath.c_str(),
            0,
            &driverRebootRequired) != FALSE;
        if (!driverPackageRemoved)
        {
            DWORD error = GetLastError();
            if (error != ERROR_FILE_NOT_FOUND && error != ERROR_NOT_FOUND)
            {
                std::wcerr << L"DiUninstallDriver failed: "
                           << ErrorMessage(error) << L" (" << error << L")\n";
            }
            else
            {
                driverPackageRemoved = true;
            }
        }

        BOOL deviceRebootRequired = FALSE;
        size_t removedCount = 0;
        bool devicesRemoved = RemoveMatchingDevices(&deviceRebootRequired, &removedCount);
        if (!driverPackageRemoved || !devicesRemoved)
        {
            return 1;
        }

        BOOL rebootRequired = driverRebootRequired || deviceRebootRequired;
        std::wcout << L"UNINSTALL_SUCCESS build=" << build
                   << L" devices=" << removedCount
                   << L" reboot=" << (rebootRequired ? L"true" : L"false") << L"\n";
        return rebootRequired ? 3010 : 0;
    }

    void PrintUsage()
    {
        std::wcout
            << L"Moshushou.DriverInstaller status\n"
            << L"Moshushou.DriverInstaller install --inf <path> [--certificate <path>]\n"
            << L"Moshushou.DriverInstaller uninstall --inf <path>\n";
    }
}

int wmain(int argc, wchar_t* argv[])
{
    _setmode(_fileno(stdout), _O_U16TEXT);
    _setmode(_fileno(stderr), _O_U16TEXT);

    if (argc < 2)
    {
        PrintUsage();
        return 5;
    }

    std::wstring command = argv[1];
    if (_wcsicmp(command.c_str(), L"status") == 0)
    {
        return PrintStatus();
    }

    bool installCommand = _wcsicmp(command.c_str(), L"install") == 0;
    bool uninstallCommand = _wcsicmp(command.c_str(), L"uninstall") == 0;
    if (!installCommand && !uninstallCommand)
    {
        PrintUsage();
        return 5;
    }

    std::filesystem::path infPath;
    std::filesystem::path certificatePath;
    for (int index = 2; index < argc; ++index)
    {
        if (_wcsicmp(argv[index], L"--inf") == 0 && index + 1 < argc)
        {
            infPath = std::filesystem::absolute(argv[++index]);
        }
        else if (installCommand && _wcsicmp(argv[index], L"--certificate") == 0 && index + 1 < argc)
        {
            certificatePath = std::filesystem::absolute(argv[++index]);
        }
        else
        {
            std::wcerr << L"Unknown or incomplete argument: " << argv[index] << L"\n";
            return 5;
        }
    }

    if (infPath.empty())
    {
        std::wcerr << L"--inf is required.\n";
        return 5;
    }
    return uninstallCommand
        ? UninstallDriver(infPath)
        : InstallDriver(infPath, certificatePath);
}
