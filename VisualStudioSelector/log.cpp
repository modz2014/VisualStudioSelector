#include <string>
#include <shobjidl_core.h>
#include <Shellapi.h>
#include <Shlwapi.h>
#include <fstream>
#include <shlobj.h>

/*void Log(const std::string& message)
{
    // Get the path to the user's desktop
    wchar_t desktopPath[MAX_PATH];
    if (SUCCEEDED(SHGetFolderPath(NULL, CSIDL_DESKTOP, NULL, 0, desktopPath)))
    {
        // Append the log file name to the desktop path
        std::wstring logFilePath = std::wstring(desktopPath) + L"\\Log.log";

        std::wofstream logFile;
        logFile.open(logFilePath, std::ios_base::app); // Append to the log file

        // Write the message to the log file
        logFile << message.c_str() << std::endl;

        logFile.close();
    }
}
*/
